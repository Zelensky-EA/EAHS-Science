/** Fast local estimates plus bounded, background supplier verification. */
function enteredUnitPrice_(value){
  if(value===undefined||value===null||String(value).trim()==='')return null;
  const price=Number(value);if(!Number.isFinite(price)||price<=0||price>100000000)throw Error('Replacement unit price must be a positive USD amount.');return price;
}
function validatePriceSource_(value){if(value&&!/^https?:\/\/[^\s<>"']+$/i.test(String(value)))throw Error('Use a full http(s) price source link.');}
// Supplier research runs in the connected ChatGPT task, not a paid API.
function automaticPricingStatus_(){return{provider:'ChatGPT',configured:null,enabled:null,managedExternally:true,itemsPerRun:5,schedule:'Managed externally'};}
function automaticPriceBatch_(defer){return{defer:!!defer,pendingPrices:[],pendingJobs:[],prices:rows_('PriceEstimates'),catalog:rows_('PriceCatalog'),jobs:rows_('PriceJobs'),configured:true};}
function planningPriceFor_(item,catalog){
  const names=new Set([nameSignature_(item.name)]);
  const match=(catalog||rows_('PriceCatalog')).filter(p=>p.kind===item.kind&&String(p.unit).toLowerCase()===String(item.unit).toLowerCase()&&(names.has(nameSignature_(p.name))||names.has(nameSignature_(catalogAlias_(p.name)))));
  return match.length===1&&Number(match[0].averageUnitCost)>0?match[0]:null;
}
function autoJobKey_(item){return priceKey_(item)+'|'+String(item.manufacturer||'').trim().toLowerCase()+'|'+String(item.catalogNumber||'').trim().toLowerCase();}
function ensureAutomaticPrice_(item,raw,batch){
  batch=batch||automaticPriceBatch_();raw=raw||{};
  const matching=batch.prices.filter(p=>priceKey_(p)===priceKey_(item));
  const valid=p=>Number.isFinite(Number(p.averageUnitCost))&&Number(p.averageUnitCost)>0;
  const existing=matching.find(p=>valid(p)&&p.matchType!=='Planning estimate')||matching.find(valid)||matching[0],entered=enteredUnitPrice_(raw.unitPrice);
  if(entered!==null){
    validatePriceSource_(raw.priceSource);
    const price={id:existing?existing.id:Utilities.getUuid(),name:item.name,kind:item.kind,unit:item.unit,averageUnitCost:entered,quotesJson:raw.priceSource?JSON.stringify([{supplier:'Staff-entered price',url:String(raw.priceSource),price:entered,packQuantity:1}]):'[]',checkedAt:Utilities.formatDate(new Date(),'America/Los_Angeles','yyyy-MM-dd'),matchType:'Entered price',notes:'USD per counted inventory unit entered by staff. Tax and shipping excluded. Confirm replacement suitability for purchase/invoice prices.',updatedBy:safe_(raw.actor||'Inventory entry',100)};
    if(existing){price._row=existing._row;if(existing._row)write_('PriceEstimates',price);Object.assign(existing,price)}else{if(batch.defer)batch.pendingPrices.push(price);else{append_('PriceEstimates',price);price._row=sheet_('PriceEstimates').getLastRow()}batch.prices.push(price)}
    finishPriceJobs_(item,'Done','Staff-entered price retained.',batch.jobs);return;
  }
  if(existing&&Number.isFinite(Number(existing.averageUnitCost))&&Number(existing.averageUnitCost)>0&&existing.matchType!=='Planning estimate')return;
  const planning=planningPriceFor_(item,batch.catalog);
  if(!existing&&planning){
    const price={id:Utilities.getUuid(),name:item.name,kind:item.kind,unit:item.unit,averageUnitCost:Number(planning.averageUnitCost),quotesJson:'[]',checkedAt:'',matchType:'Planning estimate',notes:planning.notes,updatedBy:'Automatic catalog estimate'};
    if(batch.defer)batch.pendingPrices.push(price);else{append_('PriceEstimates',price);price._row=sheet_('PriceEstimates').getLastRow()}batch.prices.push(price);
  }
  const key=autoJobKey_(item);if(batch.jobs.some(j=>j.key===key))return;
  const pack=/^(?:packs?|box|boxes)$/i.test(item.unit),job={id:Utilities.getUuid(),key,itemId:item.id,name:item.name,kind:item.kind,unit:item.unit,manufacturer:item.manufacturer||'',catalogNumber:item.catalogNumber||'',status:pack?'Needs review':batch.configured?'Queued':'API key needed',attempts:0,updatedAt:new Date().toISOString(),nextAttemptAt:'',candidatesJson:'[]',message:pack?'Enter a confirmed price per counted pack/box; its contents are unspecified.':batch.configured?'Queued for the connected ChatGPT pricing task.':'Catalog estimates work now. Waiting for the connected ChatGPT pricing task.'};
  if(batch.defer)batch.pendingJobs.push(job);else{append_('PriceJobs',job);job._row=sheet_('PriceJobs').getLastRow()}batch.jobs.push(job);
}
function finishPriceJobs_(item,status,message,jobs){(jobs||rows_('PriceJobs')).filter(j=>priceKey_(j)===priceKey_(item)).forEach(j=>{j.status=status;j.message=message;j.updatedAt=new Date().toISOString();if(j._row)write_('PriceJobs',j)});}
function initializeAutomaticPricing(){
  const data=locked_(()=>{const batch=automaticPriceBatch_(true);rows_('Items').forEach(item=>ensureAutomaticPrice_(item,{},batch));flushAutomaticPriceBatch_(batch);return getData()});
  // Disable/remove the old API schedule; the ChatGPT task has its own schedule.
  PropertiesService.getScriptProperties().setProperty('AUTOPRICING_ENABLED','false');
  ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()==='processPriceJobs').forEach(t=>ScriptApp.deleteTrigger(t));
  data.pricing=automaticPricingStatus_();data._message='Immediate estimates initialized. Supplier research is managed by the connected ChatGPT task.';return data;
}
function queuePriceSearch(raw){return locked_(()=>{
  const item=rows_('Items').find(i=>i.id===clean_(raw.itemId,100));if(!item)throw Error('Inventory item not found.');
  const batch=automaticPriceBatch_();ensureAutomaticPrice_(item,{},batch);
  if(/^(?:packs?|box|boxes)$/i.test(item.unit))throw Error('Enter a price per pack/box, or count individual pieces before requesting automatic comparison.');
  const key=autoJobKey_(item);let job=batch.jobs.find(j=>j.key===key);
  if(!job){job={id:Utilities.getUuid(),key,itemId:item.id,name:item.name,kind:item.kind,unit:item.unit,manufacturer:item.manufacturer||'',catalogNumber:item.catalogNumber||'',candidatesJson:'[]'};append_('PriceJobs',job);job._row=sheet_('PriceJobs').getLastRow()}
  if(job.status==='Running')return getData();
  Object.assign(job,{status:batch.configured?'Queued':'API key needed',attempts:0,updatedAt:new Date().toISOString(),nextAttemptAt:'',message:'Supplier verification requested. Existing staff quotes are protected.'});write_('PriceJobs',job);return getData();
});}
// Harmless compatibility handler for any remaining legacy API trigger.
function processPriceJobs(){return;}

function flushAutomaticPriceBatch_(batch){for(const [name,records] of [['PriceEstimates',batch.pendingPrices],['PriceJobs',batch.pendingJobs]]){if(records.length){const sh=sheet_(name),start=sh.getLastRow()+1;sh.getRange(start,1,records.length,TABLES[name].length).setValues(records.map(r=>TABLES[name].map(k=>r[k]==null?'':r[k])));records.forEach((r,i)=>r._row=start+i);records.length=0}}}

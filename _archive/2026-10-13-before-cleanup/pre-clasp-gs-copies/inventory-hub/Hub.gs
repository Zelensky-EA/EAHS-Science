const HUB_SHEET_ID = '1lzZHQteSyQgVgdMJbZYtrQ5hgfT6A-gUQA8Na3SYyPs';
const HUB_TEMPLATE_ID = '15l2u3t7tO4QFzQF2bzaYjgaXC1mTS0F2EDEQ4EqpRVU';
const HUB_FOLDER_ID = '1tLP3B20Z_Nrh3aevV580Pu30caXR3APe';
const HUB_TABLES = {"Updates": ["id", "title", "date", "body", "category", "status", "pinned", "expires", "updatedBy", "updatedAt"], "Meetings": ["id", "title", "date", "time", "location", "facilitator", "summary", "status", "docUrl", "updatedBy", "updatedAt"], "AgendaItems": ["id", "meetingId", "topic", "details", "submittedBy", "email", "minutes", "status", "createdAt", "updatedBy"], "Faculty": ["id", "name", "role", "courses", "room", "email", "status", "updatedBy", "updatedAt"], "Links": ["id", "title", "category", "url", "description", "status", "updatedBy", "updatedAt"], "Settings": ["key", "value"]};
const HUB_SCHEMA_CACHE_KEY_ = 'hub-schema-v1:' + HUB_SHEET_ID;
let hubExecutionBook_, hubExecutionReady_ = false;
function hubBook_(){return hubExecutionBook_ || (hubExecutionBook_ = SpreadsheetApp.openById(HUB_SHEET_ID));}
// Tab/header check runs at most once per 6 hours (run runHubSetup() from the editor after changing hub tabs).
function hubSetup_(){
  if(hubExecutionReady_)return;
  let cache=null;
  try{cache=CacheService.getScriptCache();if(cache.get(HUB_SCHEMA_CACHE_KEY_)){hubExecutionReady_=true;return}}catch(e){}
  const book=hubBook_();
  Object.keys(HUB_TABLES).forEach(name=>{
    let sh=book.getSheetByName(name);const headers=HUB_TABLES[name];
    if(!sh)sh=book.insertSheet(name);
    if(!sh.getLastRow()){sh.getRange(1,1,1,headers.length).setValues([headers]);sh.setFrozenRows(1);}
    else if(sh.getRange(1,1,1,headers.length).getValues()[0].join('|')!==headers.join('|'))throw Error('Restore the original header names and order in hub tab '+name+'.');
  });hubExecutionReady_=true;
  try{if(cache)cache.put(HUB_SCHEMA_CACHE_KEY_,'ready',21600)}catch(e){}
}
function runHubSetup(){hubExecutionReady_=false;try{CacheService.getScriptCache().remove(HUB_SCHEMA_CACHE_KEY_)}catch(e){}hubSetup_();return 'Hub tabs and headers verified.';}
function hubRows_(name){const sh=hubBook_().getSheetByName(name),count=sh.getLastRow()-1;if(count<1)return[];return sh.getRange(2,1,count,HUB_TABLES[name].length).getValues().map((r,i)=>Object.assign({_row:i+2},Object.fromEntries(HUB_TABLES[name].map((k,j)=>[k,r[j] instanceof Date?Utilities.formatDate(r[j],'America/Los_Angeles',k==='time'?'HH:mm':'yyyy-MM-dd'):r[j]])))).filter(x=>x.id||x.key);}
function getHubData(){hubSetup_();const strip=x=>{const y={...x};delete y._row;return y};return{updates:hubRows_('Updates').map(strip),meetings:hubRows_('Meetings').map(strip),agenda:hubRows_('AgendaItems').map(strip),faculty:hubRows_('Faculty').map(strip),links:hubRows_('Links').map(strip),settings:Object.fromEntries(hubRows_('Settings').map(x=>[x.key,x.value])),recordsUrl:'https://docs.google.com/spreadsheets/d/'+HUB_SHEET_ID+'/edit',templateUrl:'https://docs.google.com/document/d/'+HUB_TEMPLATE_ID+'/edit'};}
function getHubClient(){return HtmlService.createTemplateFromFile('HubClient').getRawContent();}
function hubDate_(v,required){const s=clean_(v,10);if(!s&&!required)return'';if(!/^\d{4}-\d{2}-\d{2}$/.test(s)||!Number.isFinite(Date.parse(s))||new Date(s).toISOString().slice(0,10)!==s)throw Error('Enter a valid date.');return s;}
function hubUrl_(v,required){const s=clean_(v,1200);if(!s&&!required)return'';if(!/^https?:\/\/[^\s<>"']+$/i.test(s))throw Error('Use a full http(s) link.');return s;}
function hubEmail_(v){const s=clean_(v,120);if(s&&!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(s))throw Error('Enter a valid email address.');return s;}
function saveHubRecord(raw){
  hubSetup_();const lock=LockService.getScriptLock();if(!lock.tryLock(15000))throw Error('Department records are busy. Try again.');
  try{
    const table=clean_(raw.table,40);if(!['Updates','Meetings','AgendaItems','Faculty','Links'].includes(table))throw Error('Choose a valid department record.');
    const actor=safe_(raw.actor,100);if(!actor)throw Error('Enter your name for the record.');
    const prior=raw.id?hubRows_(table).find(x=>x.id===raw.id):null;if(raw.id&&!prior)throw Error('This record is no longer available. Reload the hub.');
    const x={id:prior?prior.id:Utilities.getUuid(),updatedBy:actor,updatedAt:new Date().toISOString()};
    const text=(k,max=2000)=>safe_(raw[k],max),need=(k,max=160)=>{const v=text(k,max);if(!v)throw Error(k+' is required.');return v};
    let allowed;
    if(table==='Updates'){Object.assign(x,{title:need('title'),date:hubDate_(raw.date,true),body:need('body',4000),category:need('category',60),status:clean_(raw.status,40),pinned:raw.pinned===true||raw.pinned==='true',expires:hubDate_(raw.expires,false)});if(x.expires&&x.expires<x.date)throw Error('Expiration cannot precede the update date.');allowed=['Draft','Published','Archived'];}
    if(table==='Meetings'){
      Object.assign(x,{title:need('title'),date:hubDate_(raw.date,true),time:text('time',30),location:text('location',100),facilitator:text('facilitator',100),summary:prior?prior.summary:'',status:clean_(raw.status,40),docUrl:hubUrl_(raw.docUrl,false)});allowed=['Upcoming','Draft minutes','Approved minutes','Cancelled'];
    }
    if(table==='AgendaItems'){
      const meetingId=clean_(raw.meetingId,100),meeting=hubRows_('Meetings').find(m=>m.id===meetingId);if(meetingId&&!meeting)throw Error('Choose an existing meeting or Next meeting.');
      if(meetingId&&(!prior||meetingId!==prior.meetingId)&&(!['Upcoming','Draft minutes'].includes(meeting.status)||(meeting.status==='Upcoming'&&meeting.date<Utilities.formatDate(new Date(),'America/Los_Angeles','yyyy-MM-dd'))))throw Error('Choose an upcoming meeting or draft minutes. Approved and cancelled meetings cannot receive new topics.');
      Object.assign(x,{meetingId,topic:need('topic'),details:text('details',3000),submittedBy:prior?prior.submittedBy:actor,email:prior?prior.email:hubEmail_(raw.email),minutes:positive_(raw.minutes===undefined||raw.minutes===''?(prior?prior.minutes:5):raw.minutes,'Requested minutes',false),status:prior?clean_(raw.status,40):'Submitted',createdAt:prior?prior.createdAt:new Date().toISOString()});allowed=['Submitted','Ready for minutes','Added to minutes','Added to agenda','Deferred','Complete'];if(x.minutes>120)throw Error('Requested time cannot exceed 120 minutes.');
    }
    if(table==='Faculty'){Object.assign(x,{name:need('name',100),role:text('role',100),courses:text('courses',500),room:text('room',100),email:hubEmail_(raw.email),status:clean_(raw.status,40)});allowed=['Active','Archived'];}
    if(table==='Links'){Object.assign(x,{title:need('title'),category:need('category',80),url:hubUrl_(raw.url,true),description:text('description',1200),status:clean_(raw.status,40)});allowed=['Active','Archived'];}
    if(!allowed.includes(x.status))throw Error('Choose a valid record status.');
    const sh=hubBook_().getSheetByName(table),values=HUB_TABLES[table].map(k=>x[k]==null?'':x[k]);if(prior)sh.getRange(prior._row,1,1,values.length).setValues([values]);else sh.appendRow(values);
    let syncError='';
    if(table==='Meetings'||table==='AgendaItems')try{
      if(table==='Meetings'&&!x.docUrl&&raw.createDoc===true)hubMeetingDoc_(hubRows_('Meetings').find(m=>m.id===x.id),actor);
      if(table==='Meetings')hubSyncAgenda_(actor);
      if(table==='Meetings')hubUpdateSummary_(hubRows_('Meetings').find(m=>m.id===x.id));
    }catch(e){syncError=String(e.message||e)}
    const result=getHubData();result.message=syncError?'Record saved, but the minutes update failed: '+syncError+'. Save the meeting again to retry.':table==='AgendaItems'?'Topic submitted. Click Add to minutes when you want it included.':table==='Meetings'?'Meeting saved. Approved waiting topics have been added to the minutes.':'Department record saved.';return result;
  }finally{lock.releaseLock()}
}

function hubWriteRecord_(table,x){
  const sh=hubBook_().getSheetByName(table),values=HUB_TABLES[table].map(k=>x[k]==null?'':x[k]);
  sh.getRange(x._row,1,1,values.length).setValues([values]);
}
function hubMeetingDoc_(meeting,actor){
  if(meeting.docUrl)return;
  const copy=DriveApp.getFileById(HUB_TEMPLATE_ID).makeCopy('Science Department '+meeting.date+' '+meeting.title,DriveApp.getFolderById(HUB_FOLDER_ID));
  // Persist the link immediately so retries reuse the same document.
  meeting.docUrl='https://docs.google.com/document/d/'+copy.getId()+'/edit';
  meeting.updatedBy=actor;meeting.updatedAt=new Date().toISOString();hubWriteRecord_('Meetings',meeting);
  const doc=DocumentApp.openById(copy.getId()),body=doc.getBody();
  const values={MEETING_TITLE:meeting.title,MEETING_DATE:meeting.date,MEETING_TIME:meeting.time||'To be confirmed',MEETING_LOCATION:meeting.location||'To be confirmed',FACILITATOR:meeting.facilitator||actor};
  Object.keys(values).forEach(k=>body.replaceText('\\{\\{'+k+'\\}\\}',String(values[k])));doc.saveAndClose();
}

function hubNextMeeting_(){
  const today=Utilities.formatDate(new Date(),'America/Los_Angeles','yyyy-MM-dd');
  return hubRows_('Meetings').filter(m=>(m.status==='Upcoming'&&m.date>=today)||(m.status==='Draft minutes'&&m.date===today)).sort((a,b)=>String(a.date).localeCompare(String(b.date))||String(a.time).localeCompare(String(b.time))||String(a.id).localeCompare(String(b.id)))[0];
}
function hubDocId_(url){const match=String(url).match(/^https:\/\/docs\.google\.com\/document\/d\/([\w-]+)/);if(!match)throw Error('The meeting needs a Google Docs minutes link.');return match[1]}
function hubPutTopic_(meeting,item,actor){
  if(!['Upcoming','Draft minutes'].includes(meeting.status))throw Error('Approved or cancelled minutes cannot receive new topics.');
  hubMeetingDoc_(meeting,actor);
  const doc=DocumentApp.openById(hubDocId_(meeting.docUrl)),body=doc.getBody(),name='EAHS_AGENDA_'+item.id;
  // Named ranges identify topics without placing technical reference rows in the minutes.
  const legacy=body.getTables().some(t=>t.getNumRows()===7&&t.getRow(0).getNumCells()===2&&t.getCell(0,1).getText()==='Agenda item '+item.id);
  if(!doc.getNamedRanges(name).length&&!legacy){
    const anchor=body.getParagraphs().find(p=>p.getText().trim()==='Action items');
    let index=anchor?body.getChildIndex(anchor):body.getNumChildren();
    const heading=body.insertParagraph(index++,String(item.topic)).setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.insertParagraph(index++,'Shared by: '+item.submittedBy);
    if(item.details)body.insertParagraph(index++,'Share-out: '+item.details);
    body.insertParagraph(index++,'Notes: ');
    // Mark the whole topic including its notes; retries preserve staff-written text.
    const range=doc.newRange().addElement(heading).build();doc.addNamedRange(name,range);
  }
  doc.saveAndClose();
  item.meetingId=meeting.id;item.status='Added to minutes';item.updatedBy=actor;hubWriteRecord_('AgendaItems',item);
}
function addAgendaToMinutes(raw){
  hubSetup_();const lock=LockService.getScriptLock();if(!lock.tryLock(15000))throw Error('Department records are busy. Try again.');
  try{
    const item=hubRows_('AgendaItems').find(a=>a.id===clean_(raw.id,100));if(!item)throw Error('Reload the hub; this topic no longer exists.');
    const actor=Session.getActiveUser().getEmail()||'Department minutes';
    if(['Added to minutes','Added to agenda','Complete'].includes(item.status)){const result=getHubData();result.message='This topic is already in the meeting record.';return result}
    const meeting=hubNextMeeting_();
    if(!meeting){item.meetingId='';item.status='Ready for minutes';item.updatedBy=actor;hubWriteRecord_('AgendaItems',item);const result=getHubData();result.message='Ready for minutes. This topic will be added automatically when the next meeting is created.';return result}
    hubPutTopic_(meeting,item,actor);
    let warning='';try{hubUpdateSummary_(meeting)}catch(e){warning=' Summary refresh failed; use Refresh summary to retry.'}
    const result=getHubData();result.message='Added to '+meeting.title+' ('+meeting.date+').'+warning;return result;
  }finally{lock.releaseLock()}
}
function hubSyncAgenda_(actor){
  const next=hubNextMeeting_();if(!next)return;
  // Only topics explicitly approved using Add to minutes are picked up on meeting creation.
  hubRows_('AgendaItems').filter(a=>a.status==='Ready for minutes').forEach(a=>hubPutTopic_(next,a,actor));
}
function hubSummaryFromBody_(body){
  const lines=body.getParagraphs().map(p=>({text:p.getText().trim(),heading:p.getHeading()}));
  let section='',topics=[],notes=[],actions=[];
  const content=text=>text.replace(/_+/g,'').replace(/\s+/g,' ').trim();
  for(const p of lines){
    const t=p.text;if(/^Share outs and discussion$/i.test(t)||/^Discussion and decisions$/i.test(t)){section='discussion';continue}
    if(/^Action items$/i.test(t)){section='actions';continue}
    if(p.heading===DocumentApp.ParagraphHeading.HEADING1&&section){section='';continue}
    if(section==='discussion'){
      if(p.heading===DocumentApp.ParagraphHeading.HEADING2&&t)topics.push(content(t));
      else if(!/^Shared by:|^Share-out:|^Agenda item /i.test(t)){const value=content(t.replace(/^Notes:\s*/i,'').replace(/^Record decisions, their rationale, and any questions deferred for follow up\.\s*/i,''));if(value&&!/^Identify the decisions/i.test(value))notes.push(value)}
    }
    if(section==='actions'){
      const value=content(t.replace(/^Action:\s*/i,''));
      if(value&&!/^Owner:\s*Due:\s*$/i.test(value)&&!['Action or deliverable','Owner','Due date','Status'].includes(value))actions.push(value);
    }
  }
  // Older minutes store topic notes and actions in tables.
  body.getTables().forEach(t=>{if(t.getNumRows()===7&&t.getRow(0).getNumCells()===2&&t.getCell(0,0).getText()==='Item reference'){
    topics.push(content(t.getCell(1,1).getText()));const note=content(t.getCell(5,1).getText()),action=content(t.getCell(6,1).getText());if(note)notes.push(note);if(action)actions.push(action);
  }else if(t.getNumRows()>1&&t.getRow(0).getNumCells()===4&&t.getCell(0,0).getText()==='Action or deliverable'){
    for(let r=1;r<t.getNumRows();r++){const action=content(t.getCell(r,0).getText());if(action){const owner=content(t.getCell(r,1).getText()),due=content(t.getCell(r,2).getText());actions.push(action+(owner?' — '+owner:'')+(due?' — '+due:''))}}
  }});
  const short=(values,max)=>[...new Set(values)].slice(0,max).map(v=>v.length>180?v.slice(0,177)+'…':v).join('; ');
  return [topics.length?'Topics: '+short(topics,4)+'.':'',notes.length?'Notes: '+short(notes,3)+'.':'',actions.length?'Actions: '+short(actions,3)+'.':''].filter(Boolean).join('\n').slice(0,1200);
}
function hubUpdateSummary_(meeting){
  if(!meeting.docUrl||meeting.status==='Cancelled')return;
  const doc=DocumentApp.openById(hubDocId_(meeting.docUrl)),summary=hubSummaryFromBody_(doc.getBody());doc.saveAndClose();
  if(summary!==meeting.summary){meeting.summary=summary;hubWriteRecord_('Meetings',meeting)}
}
// Reads minutes documents WITHOUT holding the shared lock (inventory saves are never blocked),
// skips any document whose Drive "last edited" time has not changed, and locks only for the short Sheet write.
// Automatic refresh covers Upcoming and Draft-minutes meetings (5 most recent); a meeting's
// "Refresh summary" button passes its id and force=true and works for any status.
function refreshMeetingSummaries(raw){
  hubSetup_();
  const requested=clean_(raw&&raw.meetingId,100),force=!!(raw&&raw.force===true),errors=[];
  let cache=null;try{cache=CacheService.getScriptCache()}catch(e){}
  const meetings=hubRows_('Meetings').filter(m=>m.docUrl&&m.status!=='Cancelled'&&(requested?m.id===requested:['Upcoming','Draft minutes'].includes(m.status))).sort((a,b)=>String(b.date).localeCompare(String(a.date))).slice(0,requested?1:5);
  meetings.forEach(m=>{try{
    const docId=hubDocId_(m.docUrl),key='EAHS_MINUTES_STAMP_'+m.id;
    const stamp=String(DriveApp.getFileById(docId).getLastUpdated().getTime());
    if(!force&&cache&&cache.get(key)===stamp)return; // Document unchanged since the last summary.
    const summary=hubSummaryFromBody_(DocumentApp.openById(docId).getBody());
    if(summary!==m.summary){
      const lock=LockService.getScriptLock();if(!lock.tryLock(15000))throw Error('Department records are busy. Try again.');
      try{const fresh=hubRows_('Meetings').find(x=>x.id===m.id);if(fresh&&fresh.summary!==summary){fresh.summary=summary;hubWriteRecord_('Meetings',fresh)}}finally{lock.releaseLock()}
    }
    if(cache)cache.put(key,stamp,21600);
  }catch(e){errors.push(m.title+': '+String(e.message||e))}});
  const strip=x=>{const y={...x};delete y._row;return y};
  // Only meetings are returned; the page keeps its other records.
  return {meetings:hubRows_('Meetings').map(strip),message:errors.length?'Some summaries could not refresh: '+errors.join(' · '):'Summaries updated from the minutes documents.',summaryWarnings:errors};
}

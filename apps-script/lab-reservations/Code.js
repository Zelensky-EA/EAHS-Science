const CALENDAR_ID =
  'c_7aef4cc773d414d0b447c1a7ad7e6c9910c8b127ef7d22cc240b472d693f7d35@group.calendar.google.com';

const TIME_ZONE = 'America/Los_Angeles';
const ALLOWED_DOMAIN = 'salinasuhsd.org';
const NON_PRIORITY_DAYS = 14;

const PRIORITY_COURSES = [
  'AP Chemistry',
  'AP Biology',
  'Anatomy & Physiology'
];

// Add minimum-day dates using YYYY-MM-DD.
const MINIMUM_DAYS = [
  // '2026-10-14'
];

const REGULAR = [
  {
    id: 'P1',
    label: 'Period 1',
    start: '08:45',
    end: '09:44'
  },
  {
    id: 'P2',
    label: 'Period 2',
    start: '09:51',
    end: '10:48'
  },
  {
    id: 'P3',
    label: 'Period 3',
    start: '10:55',
    end: '11:52'
  },
  {
    id: 'P4',
    label: 'Period 4',
    start: '11:59',
    end: '12:56'
  },
  {
    id: 'P5',
    label: 'Period 5',
    start: '13:43',
    end: '14:40'
  },
  {
    id: 'P6',
    label: 'Period 6',
    start: '14:47',
    end: '15:44'
  }
];

const WEDNESDAY = [
  {
    id: 'B1',
    label: 'Block 1',
    start: '09:00',
    end: '10:46'
  },
  {
    id: 'B3',
    label: 'Block 3',
    start: '10:53',
    end: '12:35'
  },
  {
    id: 'B5',
    label: 'Block 5',
    start: '13:22',
    end: '15:04'
  }
];

const THURSDAY = [
  {
    id: 'B2',
    label: 'Block 2',
    start: '08:45',
    end: '10:31'
  },
  {
    id: 'B4',
    label: 'Block 4',
    start: '10:38',
    end: '12:20'
  },
  {
    id: 'B6',
    label: 'Block 6',
    start: '14:02',
    end: '15:44'
  }
];

const MINIMUM = [
  {
    id: 'P1',
    label: 'Period 1',
    start: '08:45',
    end: '09:22'
  },
  {
    id: 'P2',
    label: 'Period 2',
    start: '09:29',
    end: '10:06'
  },
  {
    id: 'P3',
    label: 'Period 3',
    start: '10:13',
    end: '10:50'
  },
  {
    id: 'P4',
    label: 'Period 4',
    start: '10:57',
    end: '11:34'
  },
  {
    id: 'P5',
    label: 'Period 5',
    start: '11:41',
    end: '12:18'
  },
  {
    id: 'P6',
    label: 'Period 6',
    start: '13:05',
    end: '13:41'
  }
];

/**
 * Receives requests from the GitHub Pages website.
 */
function doGet(e) {
  const parameters = e && e.parameter ? e.parameter : {};
  const callback = String(parameters.callback || 'callback')
    .replace(/[^a-zA-Z0-9_.$]/g, '');

  try {
    let data;

    if (parameters.action === 'availability') {
      data = getAvailability_(parameters);
    } else if (parameters.action === 'book') {
      data = book_(parameters);
    } else {
      data = {
        ok: false,
        error: 'Unknown action.'
      };
    }

    return jsonpResponse_(callback, data);
  } catch (error) {
    return jsonpResponse_(callback, {
      ok: false,
      error: error.message || String(error)
    });
  }
}

/**
 * Returns a JSONP response for the GitHub Pages website.
 */
function jsonpResponse_(callback, data) {
  return ContentService
    .createTextOutput(
      callback + '(' + JSON.stringify(data) + ')'
    )
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

/**
 * Retrieves reservations for one displayed week.
 */
function getAvailability_(parameters) {
  const start = parseDate_(parameters.start, '00:00');
  const end = parseDate_(parameters.end, '00:00');

  if ((end.getTime() - start.getTime()) / 86400000 > 8) {
    throw new Error('Only one week may be requested.');
  }

  const calendar = getCalendar_();
  const events = calendar.getEvents(start, end);

  return {
    ok: true,
    reservations: events
      .map(eventToReservation_)
      .filter(function (reservation) {
        return reservation !== null;
      })
  };
}

/**
 * Converts a calendar event into a reservation displayed by the website.
 */
function eventToReservation_(event) {
  const tag = event.getTag('labSlot');

  if (tag) {
    const parts = tag.split('|');

    return {
      date: parts[0],
      slotId: parts[1],
      teacherName: event.getTag('teacherName') || 'Teacher',
      course: event.getTag('course') || 'Reserved'
    };
  }

  // Display events created directly in Google Calendar when they overlap
  // a scheduled class period.
  const date = Utilities.formatDate(
    event.getStartTime(),
    TIME_ZONE,
    'yyyy-MM-dd'
  );

  const slot = slotsFor_(date).find(function (scheduledSlot) {
    const slotStart = parseDate_(date, scheduledSlot.start);
    const slotEnd = parseDate_(date, scheduledSlot.end);

    return (
      event.getStartTime() < slotEnd &&
      event.getEndTime() > slotStart
    );
  });

  if (!slot) {
    return null;
  }

  return {
    date: date,
    slotId: slot.id,
    teacherName: 'Calendar event',
    course: event.getTitle() || 'Reserved'
  };
}

/**
 * Validates and creates a reservation.
 */
function book_(parameters) {
  const requiredFields = [
    'date',
    'slotId',
    'teacherName',
    'teacherEmail',
    'course',
    'activity'
  ];

  requiredFields.forEach(function (field) {
    if (!parameters[field]) {
      throw new Error('Please complete every required field.');
    }
  });

  const email = String(parameters.teacherEmail)
    .trim()
    .toLowerCase();

  if (!email.endsWith('@' + ALLOWED_DOMAIN)) {
    throw new Error(
      'Please use your ' + ALLOWED_DOMAIN + ' email address.'
    );
  }

  const slots = slotsFor_(parameters.date);

  const slot = slots.find(function (scheduledSlot) {
    return scheduledSlot.id === parameters.slotId;
  });

  if (!slot) {
    throw new Error(
      'That period is not available on the selected day.'
    );
  }

  const start = parseDate_(parameters.date, slot.start);
  const end = parseDate_(parameters.date, slot.end);

  if (end < new Date()) {
    throw new Error('Past periods cannot be reserved.');
  }

  /*
   * Priority courses may reserve farther ahead.
   * All other courses are limited to 14 days.
   */
  if (PRIORITY_COURSES.indexOf(parameters.course) === -1) {
    const cutoff = new Date();

    cutoff.setDate(
      cutoff.getDate() + NON_PRIORITY_DAYS
    );

    cutoff.setHours(23, 59, 59, 999);

    if (start > cutoff) {
      throw new Error(
        'Reservations may be made up to ' +
        NON_PRIORITY_DAYS +
        ' days ahead.'
      );
    }
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const calendar = getCalendar_();
    const conflictingEvents = calendar.getEvents(start, end);

    if (conflictingEvents.length > 0) {
      throw new Error(
        'That period was just booked. Please select another slot.'
      );
    }

    const description = [
      'Teacher: ' + parameters.teacherName,
      'Email: ' + email,
      'Course: ' + parameters.course,
      'Activity: ' + parameters.activity,
      'Prep support: ' +
        (parameters.needsPrep === 'yes' ? 'Yes' : 'No'),
      parameters.notes
        ? 'Notes: ' + parameters.notes
        : ''
    ]
      .filter(Boolean)
      .join('\n');

    const event = calendar.createEvent(
      'LAB — ' +
        parameters.course +
        ' — ' +
        parameters.teacherName,
      start,
      end,
      {
        description: description,
        guests: email,
        sendInvites: true
      }
    );

    event.setTag(
      'labSlot',
      parameters.date + '|' + parameters.slotId
    );

    event.setTag(
      'teacherName',
      parameters.teacherName
    );

    event.setTag(
      'course',
      parameters.course
    );

    return {
      ok: true,
      eventId: event.getId()
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Returns the correct bell schedule for a date.
 */
function slotsFor_(dateText) {
  if (MINIMUM_DAYS.indexOf(dateText) > -1) {
    return MINIMUM;
  }

  const day = parseDate_(dateText, '12:00').getDay();

  if (day === 3) {
    return WEDNESDAY;
  }

  if (day === 4) {
    return THURSDAY;
  }

  if (day === 1 || day === 2 || day === 5) {
    return REGULAR;
  }

  return [];
}

/**
 * Converts YYYY-MM-DD and 24-hour time into a Date.
 */
function parseDate_(date, time) {
  const parts = (date + ' ' + time).match(
    /(\d{4})-(\d{2})-(\d{2}) (\d{2}):(\d{2})/
  );

  if (!parts) {
    throw new Error('Invalid date.');
  }

  return new Date(
    Number(parts[1]),
    Number(parts[2]) - 1,
    Number(parts[3]),
    Number(parts[4]),
    Number(parts[5])
  );
}

/**
 * Retrieves the configured calendar or provides a useful error.
 */
function getCalendar_() {
  const calendar = CalendarApp.getCalendarById(CALENDAR_ID);

  if (!calendar) {
    throw new Error(
      'Calendar not found. Verify the Calendar ID and deployment account.'
    );
  }

  return calendar;
}

/**
 * Optional manual test.
 */
function testCalendarAccess() {
  const calendar = CalendarApp.getCalendarById(CALENDAR_ID);

  if (calendar) {
    console.log('FOUND: ' + calendar.getName());
    console.log('ID: ' + calendar.getId());
  } else {
    console.log('NOT FOUND');
    console.log(
      'Running as: ' +
      Session.getEffectiveUser().getEmail()
    );

    CalendarApp.getAllOwnedCalendars().forEach(function (ownedCalendar) {
      console.log(
        ownedCalendar.getName() +
        ' — ' +
        ownedCalendar.getId()
      );
    });
  }
}
const msalConfig = {
  auth: {
    clientId: 'f4602006-b304-4530-8e4e-7c31c9b3cb2e',
    authority: 'https://login.microsoftonline.com/2356b269-1a6e-4033-a730-46e40484e6b5',
    redirectUri: 'https://localhost:3000/taskpane.html',
  },
  cache: {
    cacheLocation: 'localStorage',
    storeAuthStateInCookie: true,
  },
};

const loginRequest = {
  scopes: ['Calendars.ReadWrite', 'User.Read'],
};

let msalInstance;
let projectCount = 1;
const additionalEmail = 'gz.ma-abwesenheiten@ie-group.com';

document.addEventListener('DOMContentLoaded', function () {
  msalInstance = new msal.PublicClientApplication(msalConfig);

  Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
      document.getElementById('holidayForm').onsubmit = submitHoliday;
      document.getElementById('addProjectButton').onclick = addProjectFields;
      document.getElementById('removeProjectButton').onclick = removeProjectFields;
    }
  });
});

function addProjectFields() {
  projectCount++;

  const projectGroup = document.createElement('div');
  projectGroup.className = 'project-group';
  projectGroup.id = `projectGroup${projectCount}`;
  projectGroup.innerHTML = `
    <hr class="divider">
    <div class="form-group">
      <label for="projectNumber${projectCount}">Projektnummer/Funktion:</label>
      <input type="text" id="projectNumber${projectCount}" required>
    </div>
    <div class="form-group">
      <label for="projectManager${projectCount}">Projektleiter:</label>
      <input type="text" id="projectManager${projectCount}" required placeholder="Email1, Email2, ...">
    </div>
    <div class="form-group">
      <label for="projectDeputy${projectCount}">Stellvertreter des Projektes:</label>
      <input type="text" id="projectDeputy${projectCount}" required placeholder="Email1, Email2, ...">
    </div>
  `;
  document.getElementById('additionalProjects').appendChild(projectGroup);
}

function removeProjectFields() {
  if (projectCount > 1) {
    const projectGroup = document.getElementById(`projectGroup${projectCount}`);
    if (projectGroup) {
      projectGroup.remove();
      projectCount--;
    }
  } else if (projectCount === 1) {
    const initialProjectGroup = document.getElementById('initialProject');
    if (initialProjectGroup) {
      initialProjectGroup.innerHTML = '';
      projectCount--;
    }
  }
}

function submitHoliday(event) {
  event.preventDefault();

  const startDate = document.getElementById('startDate').value;
  const endDate = document.getElementById('endDate').value;
  const reason = document.getElementById('reason').value;
  const deputy = document.getElementById('deputy').value;

  const projectFields = [];
  for (let i = 1; i <= projectCount; i++) {
    const projectNumber = document.getElementById(`projectNumber${i}`);
    const projectManager = document.getElementById(`projectManager${i}`);
    const projectDeputy = document.getElementById(`projectDeputy${i}`);

    if (projectNumber && projectManager && projectDeputy) {
      projectFields.push({
        number: projectNumber.value,
        manager: projectManager.value,
        deputy: projectDeputy.value,
      });
    }
  }

  // Enddatum auf 23:59 Uhr setzen
  const endDateTime = setEndDateToEndOfDay(endDate);

  if (
    startDate &&
    endDateTime &&
    reason &&
    deputy &&
    projectFields.every((field) => field.number && field.manager && field.deputy)
  ) {
    // Formularfelder sofort zurücksetzen
    resetForm();

    msalInstance
      .loginPopup(loginRequest)
      .then((loginResponse) => {
        const account = msalInstance.getAccountByHomeId(loginResponse.account.homeAccountId);
        const accessTokenRequest = {
          scopes: ['Calendars.ReadWrite', 'User.Read'],
          account: account,
        };

        msalInstance
          .acquireTokenSilent(accessTokenRequest)
          .then((tokenResponse) => {
            const accessToken = tokenResponse.accessToken;

            getUserName(accessToken)
              .then((senderName) => {
                const subject = `${senderName}: ${reason}`;
                const bodyContent = generateBodyContent(startDate, endDate, reason, deputy, projectFields);

                const allAttendees = parseEmails(deputy).concat(
                  ...projectFields.map((field) => parseEmails(field.manager)),
                  ...projectFields.map((field) => parseEmails(field.deputy)),
                  additionalEmail
                );

                // Erstelle Ereignis für den Ersteller mit allen Teilnehmern und Status 'frei'
                createEvent(
                  startDate,
                  endDateTime,
                  subject,
                  bodyContent,
                  Office.context.mailbox.userProfile.emailAddress,
                  allAttendees,
                  accessToken,
                  'free'
                )
                  .then((eventId) => {
                    // Ändere den Status des Ereignisses auf 'beschäftigt'
                    updateEventStatus(eventId, 'busy', accessToken)
                      .then(() => {
                        showConfirmationMessage('Urlaub erfolgreich eingetragen!');
                      })
                      .catch((error) => {
                        console.error('Fehler beim Aktualisieren des Ereignisses:', error);
                        showConfirmationMessage('Fehler beim Aktualisieren des Ereignisses.');
                      });
                  })
                  .catch((error) => {
                    console.error('Fehler beim Erstellen des Ereignisses:', error);
                    showConfirmationMessage('Fehler beim Erstellen des Ereignisses.');
                  });
              })
              .catch((error) => {
                console.error('Fehler beim Abrufen des Benutzernamens:', error);
                showConfirmationMessage('Fehler beim Abrufen des Benutzernamens.');
              });
          })
          .catch((error) => {
            console.error('Fehler beim Abrufen des Zugriffstokens:', error);
            showConfirmationMessage('Fehler beim Abrufen des Zugriffstokens.');
          });
      })
      .catch((error) => {
        console.error('Fehler bei der Anmeldung:', error);
        showConfirmationMessage('Fehler bei der Anmeldung.');
      });
  } else {
    showConfirmationMessage('Bitte alle Felder ausfüllen.');
  }
}

function setEndDateToEndOfDay(endDate) {
  return `${endDate}T23:59:00`;
}

function parseEmails(emailString) {
  return emailString
    .split(',')
    .map((email) => email.trim())
    .filter((email) => isValidEmail(email));
}

function isValidEmail(email) {
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailRegex.test(email);
}

function createEvent(
  startDate,
  endDateTime,
  subject,
  bodyContent,
  organizerEmail,
  attendeesEmails,
  accessToken,
  showAs
) {
  const attendees = attendeesEmails.map((email) => ({
    emailAddress: {
      address: email,
    },
    type: 'required',
  }));

  const event = {
    subject: subject,
    start: {
      dateTime: `${startDate}T00:00:00`,
      timeZone: 'Europe/Zurich',
    },
    end: {
      dateTime: endDateTime,
      timeZone: 'Europe/Zurich',
    },
    body: {
      contentType: 'HTML',
      content: bodyContent,
    },
    showAs: showAs,
    attendees: attendees,
  };

  return fetch('https://graph.microsoft.com/v1.0/me/events', {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify(event),
  }).then((response) => {
    if (!response.ok) {
      return response.json().then((error) => {
        throw new Error(`Fehler beim Erstellen des Ereignisses für ${organizerEmail}: ${error.message}`);
      });
    }
    return response.json().then((event) => event.id);
  });
}

function updateEventStatus(eventId, showAs, accessToken) {
  const update = {
    showAs: showAs,
  };

  return fetch(`https://graph.microsoft.com/v1.0/me/events/${eventId}`, {
    method: 'PATCH',
    headers: {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify(update),
  }).then((response) => {
    if (!response.ok) {
      return response.json().then((error) => {
        throw new Error(`Fehler beim Aktualisieren des Ereignisses: ${error.message}`);
      });
    }
  });
}

function resetForm() {
  const startDateField = document.getElementById('startDate');
  const endDateField = document.getElementById('endDate');
  const reasonField = document.getElementById('reason');
  const deputyField = document.getElementById('deputy');
  const projectNumberField = document.getElementById('projectNumber1');
  const projectManagerField = document.getElementById('projectManager1');
  const projectDeputyField = document.getElementById('projectDeputy1');

  if (startDateField) startDateField.value = '';
  if (endDateField) endDateField.value = '';
  if (reasonField) reasonField.value = '';
  if (deputyField) deputyField.value = '';
  if (projectNumberField) projectNumberField.value = '';
  if (projectManagerField) projectManagerField.value = '';
  if (projectDeputyField) projectDeputyField.value = '';

  document.getElementById('additionalProjects').innerHTML = '';
  projectCount = 1;
}

function showConfirmationMessage(message) {
  const confirmationMessage = document.getElementById('confirmationMessage');
  confirmationMessage.innerText = message;
  confirmationMessage.style.display = 'block';
}

function generateBodyContent(startDate, endDate, reason, deputy, projectFields) {
  let content = `<div style="font-family: Arial; font-size: 10pt;">
                  Ferienabwesenheit von ${formatDate(startDate)} bis ${formatDate(endDate)}.<br>
                  Vorgesetzter: ${deputy}<br>
                  Grund: ${reason}<br>`;

  projectFields.forEach((field, index) => {
    content += `Projektnummer ${index + 1}: ${field.number}, Projektleiter: ${field.manager}, Projektstellvertreter: ${field.deputy}<br>`;
  });

  content += '</div>';
  return content;
}

function formatDate(dateString) {
  const date = new Date(dateString);
  const day = String(date.getDate()).padStart(2, '0');
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const year = date.getFullYear();
  return `${day}.${month}.${year}`;
}

function checkEventStatus(eventId, accessToken) {
  return fetch(`https://graph.microsoft.com/v1.0/me/events/${eventId}`, {
    method: 'GET',
    headers: {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type': 'application/json',
    },
  })
    .then((response) => {
      if (!response.ok) {
        return response.json().then((error) => {
          throw new Error(`Fehler beim Überprüfen des Ereignisses: ${error.message}`);
        });
      }
      return response.json();
    })
    .then((event) => {
      if (event.showAs === 'declined') {
        console.log('Der Stellvertreter hat den Antrag abgelehnt.');
      }
    })
    .catch((error) => {
      console.error('Fehler beim Überprüfen des Ereignisses:', error);
    });
}

function getUserName(accessToken) {
  return fetch('https://graph.microsoft.com/v1.0/me', {
    method: 'GET',
    headers: {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type': 'application/json',
    },
  }).then((response) => {
    if (!response.ok) {
      return response.json().then((error) => {
        throw new Error(`Fehler beim Abrufen des Benutzernamens: ${error.message}`);
      });
    }
    return response.json().then((user) => user.displayName);
  });
}

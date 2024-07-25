function updateGoogleForm() {

  const formId = '{id-of-form}';
  const sheetId = '{id-of-sheet}';

  const form = FormApp.openById(formId);
  const workSheet = SpreadsheetApp.openById(sheetId);

  const sheetData = workSheet.getSheets()[0].getDataRange().getValues();

  const data = prepareData(sheetData); //[ { DateInfo: Mon Jul 01 2024 17:00:00 GMT+0200 (Central European Summer Time), Persone: [ 'qui', 'quo', 'qua' ] } ]

  //console.log(questionItem.getType().toJSON(), questionItem.getId());
  console.log("Sono stati trovati " + data.length + " dal google sheet");

  console.log("Elimino tutte le domande dal form (" + form.getItems().length + ")");
  //rimuovo tutte le domande
  form.getItems().forEach(item => form.deleteItem(item));

  console.log("le vecchie domande sono state rimosse");

  console.log("Creo domanda per selezionare la timezone");

  let timezones = [
    {label: "Niue Time, Samoa Time", minutesDiff: -660, shortString: "GMT-11:00"},
    {label: "Hawaii-Aleutian Time, Tahiti Time, Cook Islands Time", minutesDiff: -600, shortString: "GMT-10:00"},
    {label: "Marquesas Time", minutesDiff: -570, shortString: "GMT-09:30"},
    {label: "Gambier Time, Hawaii-Aleutian Time (Adak)", minutesDiff: -540, shortString: "GMT-09:00"},
    {label: "Alaska Time, Pitcairn Time", minutesDiff: -480, shortString: "GMT-08:00"},
    {label: "Pacific Time, Mexican Pacific Time, Yukon Time", minutesDiff: -420, shortString: "GMT-07:00"},
    {label: "Mountain Time, Galapagos Time, Easter Island Time", minutesDiff: -360, shortString: "GMT-06:00"},
    {label: "Central Time, Acre Time, Colombia Time, Eastern Time, Ecuador Time, Peru Standard Time", minutesDiff: -300, shortString: "GMT-05:00"},
    {label: "Atlantic Time, Eastern Time, Amazon Time, Bolivia Time, Chile Time, Cuba Time, Guyana Time, Paraguay Time, Venezuela Time", minutesDiff: -240, shortString: "GMT-04:00"},
    {label: "Argentina Time, Brasilia Time,Chile Time, Falkland Islands Time, French Guiana Time, Paraguay Time, Rothera Time, Suriname Time, Uruguay Time, West Greenland Time", minutesDiff: -180, shortString: "GMT-03:00"},
    {label: "South Georgia Time", minutesDiff: -120, shortString: "GMT-02:00"},
    {label: "Azores Time", minutesDiff: -60, shortString: "GMT-01:00"},
    {label: "Universal Time Coordinated, Greenwich Time", minutesDiff: 0, shortString: "GMT+00:00"},
    {label: "Western European Time, United Kingdom Time, Ireland Time, Morocco Time, West Africa Standard Time", minutesDiff: 60, shortString: "GMT+01:00"},
    {label: "Central European Time, Central Africa Time, South Africa Time, Troll Time", minutesDiff: 120, shortString: "GMT+02:00"},
    {label: "Eastern European Time, Arabian Time, East Africa Time, Famagusta Time, Israel Time, Jordan Time, Kirov Time, Moscow Time, Syria Time, Türkiye Time, Volgograd Time", minutesDiff: 180, shortString: "GMT+03:00"},
    {label: "Iran Time", minutesDiff: 210, shortString: "GMT+03:30"},
    {label: "Armenia Time, Astrakhan Time, Azerbaijan Time, Georgia Time, Gulf Time, Mauritius Time, Samara Time, Saratov Time, Ulyanovsk Time", minutesDiff: 240, shortString: "GMT+04:00"},
    {label: "Afghanistan Time", minutesDiff: 270, shortString: "GMT+04:30"},
    {label: "Pakistan Time, Maldives Time, Mawson Time, Tajikistan Time, Turkmenistan Time, Uzbekistan Time, Vostok Time, West Kazakhstan Time, Yekaterinburg Time", minutesDiff: 300, shortString: "GMT+05:00"},
    {label: "India Standard Time", minutesDiff: 330, shortString: "GMT+05:30"},
    {label: "Nepal Time", minutesDiff: 345, shortString: "GMT+05:45"},
    {label: "Bangladesh Time, Bhutan Time, Indian Ocean Time, Kyrgyzstan Time, Omsk Time, Urumqi Time", minutesDiff: 360, shortString: "GMT+06:00"},
    {label: "Myanmar Time", minutesDiff: 390, shortString: "GMT+06:30"},
    {label: "Barnaul Time, Indochina Time, Davis Time, Hovd Time, Indochina Time, Krasnoyarsk Time, Novosibirsk Time, Tomsk Time, Western Indonesia Time", minutesDiff: 420, shortString: "GMT+07:00"},
    {label: "Singapore Time, Philippine Time, Australian Western Time, Malaysia Time, Central Indonesia Time, China Time, Hong Kong Time, Irkutsk Time, Taipei Time, Ulaanbaatar Standard Time, Ulaanbaatar Time", minutesDiff: 480, shortString: "GMT+08:00"},
    {label: "Australian Central Western Time", minutesDiff: 525, shortString: "GMT+08:45"},
    {label: "Japan Time, Korean Time, Eastern Indonesia Time, East Timor Time, Palau Time, Yakutsk Time", minutesDiff: 540, shortString: "GMT+09:00"},
    {label: "Central Australia Time", minutesDiff: 570, shortString: "GMT+09:30"},
    {label: "Eastern Australia Time, ChamorroTime, Papua New Guinea Time, Vladivostok Standard Time", minutesDiff: 600, shortString: "GMT+10:00"},
    {label: "Lord Howe Time", minutesDiff: 630, shortString: "GMT+10:30"},
    {label: "Bougainville Time, Kosrae Time, Magadan Time, New Caledonia Time, Norfolk Island Time, Solomon Islands Time, Sakhalin Time, Srednekolymsk Time, Vanuatu Time", minutesDiff: 660, shortString: "GMT+11:00"},
    {label: "New Zealand Time, Fiji Time, Anadyr Time, Gilbert Islands Time, Marshall Islands Time, Nauru Time, Petropavlovsk-Kamchatski Time", minutesDiff: 720, shortString: "GMT+12:00"},
    {label: "Chatham Time", minutesDiff: 765, shortString: "GMT+12:45"},
    {label: "Apia Time, Phoenix Islands Time, Tokelau Time, Tonga Time", minutesDiff: 780, shortString: "GMT+13:00"},
    {label: "Line Islands Time", minutesDiff: 840, shortString: "GMT+14:00"}
  ];


  //per test ne prendo solo 5
  // timezones = timezones.slice(0, 4);
  // timezones.push({label: "Central European Time – Rome", minutesDiff: 120, shortString: "GMT+02:00"});
  const listItem = form.addListItem();
  listItem.setTitle("What time zone are you in?");
  sleep(1000);
  listItem.setChoiceValues(timezones.map(tz => tz.shortString + " - " + tz.label));

  console.log("Creata domanda per selezionare la timezone");

  const timezoneToPgBreakMap = new Map();


  let uniqueTimezones = [...new Map(timezones.map(item => [item.minutesDiff, item])).values()];
  //creo una nuova sezione per ogni timezones
  for (const timezone of uniqueTimezones) {
    const mainTitle = timezone.shortString;
    const pgBreakItem = form.addPageBreakItem()
    pgBreakItem.setTitle(mainTitle);
    pgBreakItem.setHelpText(
      "Please select all the times (at which the focus group will begin) in which you are available.\n" +
      "Keep in mind that the focus group will last between 1 hour and 2 hours. \n" + "All times are shown in your selected time zone."
    );
    pgBreakItem.setGoToPage(FormApp.PageNavigationType.SUBMIT);

    // save this page break to the map
    timezoneToPgBreakMap.set(mainTitle, pgBreakItem);

    const jsTz = getCurrentAdjTime();

    const refTimezoneStr = "GMT+02:00";
    const refTimezone = timezones.find(tz => tz.shortString === refTimezoneStr);
    //va convertita la data nel fuso orario corrente
    const localDates = data.map(d => new Date(d.DateInfo.getTime() + timezone.minutesDiff * 60000 - refTimezone.minutesDiff * 60000));
    //tolgo l'ora alla data
    const dateOnlyDistinct = [...new Set(localDates.map(d => d.toLocaleDateString('en-US')))];
    for (const date of dateOnlyDistinct) {
      //Aggiungo una casella di controllo per chiedere la disponibilità
      let dateFormatted = new Date(date);
      let options = {weekday: 'long', year: 'numeric', month: 'long', day: 'numeric'};
      let dateString = dateFormatted.toLocaleDateString("en-US", options);
      const choices = localDates
        .filter(d => d.toLocaleDateString('en-US') === date)
        .map(d => d.toLocaleTimeString('en-US', {hour: '2-digit', minute: '2-digit'}));
      choices.push("I'm not available in any of this times");
      console.log("Creo domanda per la tz " + timezone.shortString + " e la data " + dateString + " con " + choices.length + " opzioni");
      const checkItem = form.addCheckboxItem();
      sleep(500);
      checkItem.setTitle("On " + dateString + ", I am available to participate in a focus group beginning at: ");
      sleep(1000);
      checkItem.setChoiceValues(choices);
      console.log("Domanda creata");
    }


  }

  console.log("Creo associazioni tra le domande e le pagine");
  // Crea una lista di scelte con le rispettive opzioni di navigazione
  const choicesWithNavigation = timezones.map(tz => {
    const title = tz.shortString + " - " + tz.label;
    const page = timezoneToPgBreakMap.get(tz.shortString);
    sleep(300);
    return listItem.createChoice(title, page);
  });


  listItem.setChoices(choicesWithNavigation);
  console.log("Associazioni create");
}

function sleep(ms) {
  Utilities.sleep(ms);
}


function prepareData(sheetData) {
  const kdDate = 'Data disponibilità';
  const kdTimeFrom = 'Ora disponibilità from';
  const keyPersone = 'Persone';
  const data = [];

  const headers = sheetData.shift().map(header => header.trim());
  const dateIndex = headers.indexOf(kdDate);
  const timeFromIndex = headers.indexOf(kdTimeFrom);
  const personIndex = headers.indexOf(keyPersone);
  sheetData.forEach(row => {
    if (row[dateIndex] === '')
      return;
    const dt = new Date(row[dateIndex]);
    const time = row[timeFromIndex].split(':');
    const DateInfo = new Date(dt.getFullYear(), dt.getMonth(), dt.getDate(), time[0], time[1]);
    const Persone = row[personIndex].split(',');
    data.push({DateInfo, Persone});
  });

  return data;
}

function getCurrentAdjTime() {
  const match = new RegExp(/GMT([+-]\d{4})/).exec(new Date().toString());
  if (match !== null) {
    const offset = match[1];
    const sign = offset[0];
    const hours = parseInt(offset.substring(1, 3));
    const minutes = parseInt(offset.substring(3, 4));
    const totalMinutes = hours * 60 + minutes;
    const finalOffset = sign === '+' ? totalMinutes : -totalMinutes;
    return finalOffset;
  }
}

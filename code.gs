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
    {label: "International Date Line West", minutesDiff: -720, shortString: "UTC-12:00"},
    {label: "Coordinated Universal Time-11", minutesDiff: -660, shortString: "UTC-11:00"},
    {label: "Hawaii", minutesDiff: -600, shortString: "UTC-10:00"},
    {label: "Alaska", minutesDiff: -540, shortString: "UTC-09:00"},
    {label: "Baja California, Pacific Time (US and Canada)", minutesDiff: -480, shortString: "UTC-08:00"},
    {label: "Chihuahua, La Paz, Mazatlan, Arizona, Mountain Time (US and Canada)", minutesDiff: -420, shortString: "UTC-07:00"},
    {label: "Central America, Central Time (US and Canada), Saskatchewan, Guadalajara, Mexico City, Monterey", minutesDiff: -360, shortString: "UTC-06:00"},
    {label: "Bogota, Lima, Quito, Indiana (East), Eastern Time (US and Canada)", minutesDiff: -300, shortString: "UTC-05:00"},
    {label: "Caracas", minutesDiff: -270, shortString: "UTC-04:30"},
    {label: "Atlantic Time (Canada), Asuncion, Georgetown, La Paz, Manaus, San Juan, Cuiaba, Santiago", minutesDiff: -240, shortString: "UTC-04:00"},
    {label: "Newfoundland", minutesDiff: -210, shortString: "UTC-03:30"},
    {label: "Brasilia, Greenland, Cayenne, Fortaleza, Buenos Aires, Montevideo", minutesDiff: -180, shortString: "UTC-03:00"},
    {label: "Coordinated Universal Time-2", minutesDiff: -120, shortString: "UTC-02:00"},
    {label: "Cape Verde, Azores", minutesDiff: -60, shortString: "UTC-01:00"},
    {label: "Casablanca, Monrovia, Reykjavik, Dublin, Edinburgh, Lisbon, London, Coordinated Universal Time", minutesDiff: 0, shortString: "UTC+00:00"},
    {
      label: "Amsterdam, Berlin, Bern, Rome, Stockholm, Vienna, Brussels, Copenhagen, Madrid, Paris, West Central Africa, Belgrade, Bratislava, Budapest, Ljubljana, Prague, Sarajevo, Skopje, Warsaw, Zagreb, Windhoek",
      minutesDiff: 60,
      shortString: "UTC+01:00"
    },
    {
      label: "Athens, Bucharest, Istanbul, Helsinki, Kyiv, Riga, Sofia, Tallinn, Vilnius, Cairo, Damascus, Amman, Harare, Pretoria, Jerusalem, Beirut",
      minutesDiff: 120,
      shortString: "UTC+02:00"
    },
    {label: "Baghdad, Minsk, Kuwait, Riyadh, Nairobi", minutesDiff: 180, shortString: "UTC+03:00"},
    {label: "Tehran", minutesDiff: 210, shortString: "UTC+03:30"},
    {label: "Moscow, St. Petersburg, Volgograd, Tbilisi, Yerevan, Abu Dhabi, Muscat, Baku, Port Louis", minutesDiff: 240, shortString: "UTC+04:00"},
    {label: "Kabul", minutesDiff: 270, shortString: "UTC+04:30"},
    {label: "Tashkent, Islamabad, Karachi", minutesDiff: 300, shortString: "UTC+05:00"},
    {label: "Sri Jayewardenepura Kotte, Chennai, Kolkata, Mumbai, New Delhi", minutesDiff: 330, shortString: "UTC+05:30"},
    {label: "Kathmandu", minutesDiff: 345, shortString: "UTC+05:45"},
    {label: "Astana, Dhaka, Yekaterinburg", minutesDiff: 360, shortString: "UTC+06:00"},
    {label: "Yangon", minutesDiff: 390, shortString: "UTC+06:30"},
    {label: "Bangkok, Hanoi, Jakarta, Novosibirsk", minutesDiff: 420, shortString: "UTC+07:00"},
    {label: "Krasnoyarsk, Ulaanbaatar, Beijing, Chongqing, Hong Kong, Urumqi, Perth, Kuala Lumpur, Singapore, Taipei", minutesDiff: 480, shortString: "UTC+08:00"},
    {label: "Irkutsk, Seoul, Osaka, Sapporo, Tokyo", minutesDiff: 540, shortString: "UTC+09:00"},
    {label: "Darwin, Adelaide", minutesDiff: 570, shortString: "UTC+09:30"},
    {label: "Hobart, Yakutsk, Brisbane, Guam, Port Moresby, Canberra, Melbourne, Sydney", minutesDiff: 600, shortString: "UTC+10:00"},
    {label: "Vladivostok, Solomon Islands, New Caledonia", minutesDiff: 660, shortString: "UTC+11:00"},
    {label: "Coordinated Universal Time+12, Fiji, Marshall Islands, Magadan, Auckland, Wellington", minutesDiff: 720, shortString: "UTC+12:00"},
    {label: "Nuku'alofa, Samoa", minutesDiff: 780, shortString: "UTC+13:00"}
  ];


  //per test ne prendo solo 5
  //timezones = timezones.slice(0, 1);
  const listItem = form.addListItem();
  listItem.setTitle("What time zone are you in?");
  sleep(1000);
  listItem.setChoiceValues(timezones.map(tz => tz.shortString+" - "+tz.label));

  console.log("Creata domanda per selezionare la timezone");

  const timezoneToPgBreakMap = new Map();

  //creo una nuova sezione per ogni timezones
  for (const timezone of timezones) {
    const mainTitle = timezone.shortString + " - " + timezone.label;
    const pgBreakItem = form.addPageBreakItem()
    pgBreakItem.setTitle(mainTitle);
    pgBreakItem.setHelpText(
      "Please select the times (at which the focus group will begin) in which you are available.\n" +
      "Keep in mind that the focus group will last between 1 hour and 2 hours."
    );
    pgBreakItem.setGoToPage(FormApp.PageNavigationType.SUBMIT);

    // save this page break to the map
    timezoneToPgBreakMap.set(mainTitle, pgBreakItem);

    const jsTz = getCurrentAdjTime();

    //va convertita la data nel fuso orario corrente
    const localDates = data.map(d => new Date(d.DateInfo.getTime() + timezone.minutesDiff * 60000 + jsTz * 60000));
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
      console.log("Creo domanda per la tz " + timezone.shortString + " e la data " + dateString + " con " + choices.length + " opzioni");
      const checkItem = form.addCheckboxItem();
      sleep(500);
      checkItem.setTitle("When would you be available to participate in a focus group on " + dateString + "? (you can select multiple options) [" + timezone.shortString + "]");
      sleep(1000);
      checkItem.setChoiceValues(choices);
      console.log("Domanda creata");
    }


  }

  console.log("Creo associazioni tra le domande e le pagine");
  // Crea una lista di scelte con le rispettive opzioni di navigazione
  const choicesWithNavigation = timezones.map(tz => {
    const title = tz.shortString + " - " + tz.label;
    const page = timezoneToPgBreakMap.get(title);
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

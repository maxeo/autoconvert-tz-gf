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
    {label: "(GMT-11:00) Niue Time – Niue", minutesDiff: -660, shortString: "GMT-11:00"},
    {label: "(GMT-11:00) Samoa Standard Time – Pago Pago", minutesDiff: -660, shortString: "GMT-11:00"},
    {label: "(GMT-10:00) Cook Islands Standard Time – Rarotonga", minutesDiff: -600, shortString: "GMT-10:00"},
    {label: "(GMT-10:00) Hawaii-Aleutian Standard Time – Honolulu", minutesDiff: -600, shortString: "GMT-10:00"},
    {label: "(GMT-10:00) Tahiti Time – Tahiti", minutesDiff: -600, shortString: "GMT-10:00"},
    {label: "(GMT-09:30) Marquesas Time – Marquesas", minutesDiff: -570, shortString: "GMT-09:30"},
    {label: "(GMT-09:00) Gambier Time – Gambier", minutesDiff: -540, shortString: "GMT-09:00"},
    {label: "(GMT-09:00) Hawaii-Aleutian Time (Adak) – Adak", minutesDiff: -540, shortString: "GMT-09:00"},
    {label: "(GMT-08:00) Alaska Time – Anchorage", minutesDiff: -480, shortString: "GMT-08:00"},
    {label: "(GMT-08:00) Alaska Time – Juneau", minutesDiff: -480, shortString: "GMT-08:00"},
    {label: "(GMT-08:00) Alaska Time – Metlakatla", minutesDiff: -480, shortString: "GMT-08:00"},
    {label: "(GMT-08:00) Alaska Time – Nome", minutesDiff: -480, shortString: "GMT-08:00"},
    {label: "(GMT-08:00) Alaska Time – Sitka", minutesDiff: -480, shortString: "GMT-08:00"},
    {label: "(GMT-08:00) Alaska Time – Yakutat", minutesDiff: -480, shortString: "GMT-08:00"},
    {label: "(GMT-08:00) Pitcairn Time – Pitcairn", minutesDiff: -480, shortString: "GMT-08:00"},
    {label: "(GMT-07:00) Mexican Pacific Standard Time – Hermosillo", minutesDiff: -420, shortString: "GMT-07:00"},
    {label: "(GMT-07:00) Mexican Pacific Standard Time – Mazatlan", minutesDiff: -420, shortString: "GMT-07:00"},
    {label: "(GMT-07:00) Mountain Standard Time – Phoenix", minutesDiff: -420, shortString: "GMT-07:00"},
    {label: "(GMT-07:00) Mountain Standard Time – Dawson Creek", minutesDiff: -420, shortString: "GMT-07:00"},
    {label: "(GMT-07:00) Mountain Standard Time – Fort Nelson", minutesDiff: -420, shortString: "GMT-07:00"},
    {label: "(GMT-07:00) Pacific Time – Los Angeles", minutesDiff: -420, shortString: "GMT-07:00"},
    {label: "(GMT-07:00) Pacific Time – Tijuana", minutesDiff: -420, shortString: "GMT-07:00"},
    {label: "(GMT-07:00) Pacific Time – Vancouver", minutesDiff: -420, shortString: "GMT-07:00"},
    {label: "(GMT-07:00) Yukon Time – Dawson", minutesDiff: -420, shortString: "GMT-07:00"},
    {label: "(GMT-07:00) Yukon Time – Whitehorse", minutesDiff: -420, shortString: "GMT-07:00"},
    {label: "(GMT-06:00) Central Standard Time – Bahía de Banderas", minutesDiff: -360, shortString: "GMT-06:00"},
    {label: "(GMT-06:00) Central Standard Time – Belize", minutesDiff: -360, shortString: "GMT-06:00"},
    {label: "(GMT-06:00) Central Standard Time – Chihuahua", minutesDiff: -360, shortString: "GMT-06:00"},
    {label: "(GMT-06:00) Central Standard Time – Costa Rica", minutesDiff: -360, shortString: "GMT-06:00"},
    {label: "(GMT-06:00) Central Standard Time – El Salvador", minutesDiff: -360, shortString: "GMT-06:00"},
    {label: "(GMT-06:00) Central Standard Time – Guatemala", minutesDiff: -360, shortString: "GMT-06:00"},
    {label: "(GMT-06:00) Central Standard Time – Managua", minutesDiff: -360, shortString: "GMT-06:00"},
    {label: "(GMT-06:00) Central Standard Time – Mexico City", minutesDiff: -360, shortString: "GMT-06:00"},
    {label: "(GMT-06:00) Central Standard Time – Monterrey", minutesDiff: -360, shortString: "GMT-06:00"},
    {label: "(GMT-06:00) Central Standard Time – Mérida", minutesDiff: -360, shortString: "GMT-06:00"},
    {label: "(GMT-06:00) Central Standard Time – Regina", minutesDiff: -360, shortString: "GMT-06:00"},
    {label: "(GMT-06:00) Central Standard Time – Swift Current", minutesDiff: -360, shortString: "GMT-06:00"},
    {label: "(GMT-06:00) Central Standard Time – Tegucigalpa", minutesDiff: -360, shortString: "GMT-06:00"},
    {label: "(GMT-06:00) Easter Island Time – Easter", minutesDiff: -360, shortString: "GMT-06:00"},
    {label: "(GMT-06:00) Galapagos Time – Galapagos", minutesDiff: -360, shortString: "GMT-06:00"},
    {label: "(GMT-06:00) Mountain Time – Boise", minutesDiff: -360, shortString: "GMT-06:00"},
    {label: "(GMT-06:00) Mountain Time – Cambridge Bay", minutesDiff: -360, shortString: "GMT-06:00"},
    {label: "(GMT-06:00) Mountain Time – Ciudad Juárez", minutesDiff: -360, shortString: "GMT-06:00"},
    {label: "(GMT-06:00) Mountain Time – Denver", minutesDiff: -360, shortString: "GMT-06:00"},
    {label: "(GMT-06:00) Mountain Time – Edmonton", minutesDiff: -360, shortString: "GMT-06:00"},
    {label: "(GMT-06:00) Mountain Time – Inuvik", minutesDiff: -360, shortString: "GMT-06:00"},
    {label: "(GMT-05:00) Acre Standard Time – Eirunepe", minutesDiff: -300, shortString: "GMT-05:00"},
    {label: "(GMT-05:00) Acre Standard Time – Rio Branco", minutesDiff: -300, shortString: "GMT-05:00"},
    {label: "(GMT-05:00) Central Time – Beulah, North Dakota", minutesDiff: -300, shortString: "GMT-05:00"},
    {label: "(GMT-05:00) Central Time – Center, North Dakota", minutesDiff: -300, shortString: "GMT-05:00"},
    {label: "(GMT-05:00) Central Time – Chicago", minutesDiff: -300, shortString: "GMT-05:00"},
    {label: "(GMT-05:00) Central Time – Knox, Indiana", minutesDiff: -300, shortString: "GMT-05:00"},
    {label: "(GMT-05:00) Central Time – Matamoros", minutesDiff: -300, shortString: "GMT-05:00"},
    {label: "(GMT-05:00) Central Time – Menominee", minutesDiff: -300, shortString: "GMT-05:00"},
    {label: "(GMT-05:00) Central Time – New Salem, North Dakota", minutesDiff: -300, shortString: "GMT-05:00"},
    {label: "(GMT-05:00) Central Time – Ojinaga", minutesDiff: -300, shortString: "GMT-05:00"},
    {label: "(GMT-05:00) Central Time – Rankin Inlet", minutesDiff: -300, shortString: "GMT-05:00"},
    {label: "(GMT-05:00) Central Time – Resolute", minutesDiff: -300, shortString: "GMT-05:00"},
    {label: "(GMT-05:00) Central Time – Tell City, Indiana", minutesDiff: -300, shortString: "GMT-05:00"},
    {label: "(GMT-05:00) Central Time – Winnipeg", minutesDiff: -300, shortString: "GMT-05:00"},
    {label: "(GMT-05:00) Colombia Standard Time – Bogota", minutesDiff: -300, shortString: "GMT-05:00"},
    {label: "(GMT-05:00) Eastern Standard Time – Panama", minutesDiff: -300, shortString: "GMT-05:00"},
    {label: "(GMT-05:00) Eastern Standard Time – Cancún", minutesDiff: -300, shortString: "GMT-05:00"},
    {label: "(GMT-05:00) Eastern Standard Time – Jamaica", minutesDiff: -300, shortString: "GMT-05:00"},
    {label: "(GMT-05:00) Ecuador Time – Guayaquil", minutesDiff: -300, shortString: "GMT-05:00"},
    {label: "(GMT-05:00) Peru Standard Time – Lima", minutesDiff: -300, shortString: "GMT-05:00"},
    {label: "(GMT-04:00) Amazon Standard Time – Boa Vista", minutesDiff: -240, shortString: "GMT-04:00"},
    {label: "(GMT-04:00) Amazon Standard Time – Campo Grande", minutesDiff: -240, shortString: "GMT-04:00"},
    {label: "(GMT-04:00) Amazon Standard Time – Cuiaba", minutesDiff: -240, shortString: "GMT-04:00"},
    {label: "(GMT-04:00) Amazon Standard Time – Manaus", minutesDiff: -240, shortString: "GMT-04:00"},
    {label: "(GMT-04:00) Amazon Standard Time – Porto Velho", minutesDiff: -240, shortString: "GMT-04:00"},
    {label: "(GMT-04:00) Atlantic Standard Time – Puerto Rico", minutesDiff: -240, shortString: "GMT-04:00"},
    {label: "(GMT-04:00) Atlantic Standard Time – Barbados", minutesDiff: -240, shortString: "GMT-04:00"},
    {label: "(GMT-04:00) Atlantic Standard Time – Martinique", minutesDiff: -240, shortString: "GMT-04:00"},
    {label: "(GMT-04:00) Atlantic Standard Time – Santo Domingo", minutesDiff: -240, shortString: "GMT-04:00"},
    {label: "(GMT-04:00) Bolivia Time – La Paz", minutesDiff: -240, shortString: "GMT-04:00"},
    {label: "(GMT-04:00) Chile Time – Santiago", minutesDiff: -240, shortString: "GMT-04:00"},
    {label: "(GMT-04:00) Cuba Time – Havana", minutesDiff: -240, shortString: "GMT-04:00"},
    {label: "(GMT-04:00) Eastern Time – Detroit", minutesDiff: -240, shortString: "GMT-04:00"},
    {label: "(GMT-04:00) Eastern Time – Grand Turk", minutesDiff: -240, shortString: "GMT-04:00"},
    {label: "(GMT-04:00) Eastern Time – Indianapolis", minutesDiff: -240, shortString: "GMT-04:00"},
    {label: "(GMT-04:00) Eastern Time – Iqaluit", minutesDiff: -240, shortString: "GMT-04:00"},
    {label: "(GMT-04:00) Eastern Time – Louisville", minutesDiff: -240, shortString: "GMT-04:00"},
    {label: "(GMT-04:00) Eastern Time – Marengo, Indiana", minutesDiff: -240, shortString: "GMT-04:00"},
    {label: "(GMT-04:00) Eastern Time – Monticello, Kentucky", minutesDiff: -240, shortString: "GMT-04:00"},
    {label: "(GMT-04:00) Eastern Time – Toronto", minutesDiff: -240, shortString: "GMT-04:00"},
    {label: "(GMT-04:00) Eastern Time – New York", minutesDiff: -240, shortString: "GMT-04:00"},
    {label: "(GMT-04:00) Eastern Time – Petersburg, Indiana", minutesDiff: -240, shortString: "GMT-04:00"},
    {label: "(GMT-04:00) Eastern Time – Port-au-Prince", minutesDiff: -240, shortString: "GMT-04:00"},
    {label: "(GMT-04:00) Eastern Time – Vevay, Indiana", minutesDiff: -240, shortString: "GMT-04:00"},
    {label: "(GMT-04:00) Eastern Time – Vincennes, Indiana", minutesDiff: -240, shortString: "GMT-04:00"},
    {label: "(GMT-04:00) Eastern Time – Winamac, Indiana", minutesDiff: -240, shortString: "GMT-04:00"},
    {label: "(GMT-04:00) Guyana Time – Guyana", minutesDiff: -240, shortString: "GMT-04:00"},
    {label: "(GMT-04:00) Paraguay Time – Asunción", minutesDiff: -240, shortString: "GMT-04:00"},
    {label: "(GMT-04:00) Venezuela Time – Caracas", minutesDiff: -240, shortString: "GMT-04:00"},
    {label: "(GMT-03:00) Argentina Standard Time – Buenos Aires", minutesDiff: -180, shortString: "GMT-03:00"},
    {label: "(GMT-03:00) Argentina Standard Time – Catamarca", minutesDiff: -180, shortString: "GMT-03:00"},
    {label: "(GMT-03:00) Argentina Standard Time – Cordoba", minutesDiff: -180, shortString: "GMT-03:00"},
    {label: "(GMT-03:00) Argentina Standard Time – Jujuy", minutesDiff: -180, shortString: "GMT-03:00"},
    {label: "(GMT-03:00) Argentina Standard Time – La Rioja", minutesDiff: -180, shortString: "GMT-03:00"},
    {label: "(GMT-03:00) Argentina Standard Time – Mendoza", minutesDiff: -180, shortString: "GMT-03:00"},
    {label: "(GMT-03:00) Argentina Standard Time – Rio Gallegos", minutesDiff: -180, shortString: "GMT-03:00"},
    {label: "(GMT-03:00) Argentina Standard Time – Salta", minutesDiff: -180, shortString: "GMT-03:00"},
    {label: "(GMT-03:00) Argentina Standard Time – San Juan", minutesDiff: -180, shortString: "GMT-03:00"},
    {label: "(GMT-03:00) Argentina Standard Time – San Luis", minutesDiff: -180, shortString: "GMT-03:00"},
    {label: "(GMT-03:00) Argentina Standard Time – Tucuman", minutesDiff: -180, shortString: "GMT-03:00"},
    {label: "(GMT-03:00) Argentina Standard Time – Ushuaia", minutesDiff: -180, shortString: "GMT-03:00"},
    {label: "(GMT-03:00) Brasilia Standard Time – Araguaina", minutesDiff: -180, shortString: "GMT-03:00"},
    {label: "(GMT-03:00) Brasilia Standard Time – Belem", minutesDiff: -180, shortString: "GMT-03:00"},
    {label: "(GMT-03:00) Brasilia Standard Time – Belo Horizonte", minutesDiff: -180, shortString: "GMT-03:00"},
    {label: "(GMT-03:00) Brasilia Standard Time – Campo Grande", minutesDiff: -180, shortString: "GMT-03:00"},
    {label: "(GMT-03:00) Brasilia Standard Time – Cuiaba", minutesDiff: -180, shortString: "GMT-03:00"},
    {label: "(GMT-03:00) Brasilia Standard Time – Eirunepe", minutesDiff: -180, shortString: "GMT-03:00"},
    {label: "(GMT-03:00) Brasilia Standard Time – Fortaleza", minutesDiff: -180, shortString: "GMT-03:00"},
    {label: "(GMT-03:00) Brasilia Standard Time – Macapa", minutesDiff: -180, shortString: "GMT-03:00"},
    {label: "(GMT-03:00) Brasilia Standard Time – Manaus", minutesDiff: -180, shortString: "GMT-03:00"},
    {label: "(GMT-03:00) Brasilia Standard Time – Maceio", minutesDiff: -180, shortString: "GMT-03:00"},
    {label: "(GMT-03:00) Brasilia Standard Time – Porto Velho", minutesDiff: -180, shortString: "GMT-03:00"},
    {label: "(GMT-03:00) Brasilia Standard Time – Recife", minutesDiff: -180, shortString: "GMT-03:00"},
    {label: "(GMT-03:00) Brasilia Standard Time – Rio Branco", minutesDiff: -180, shortString: "GMT-03:00"},
    {label: "(GMT-03:00) Brasilia Standard Time – Santarem", minutesDiff: -180, shortString: "GMT-03:00"},
    {label: "(GMT-03:00) Brasilia Standard Time – Sao Paulo", minutesDiff: -180, shortString: "GMT-03:00"},
    {label: "(GMT-03:00) Chile Standard Time – Punta Arenas", minutesDiff: -180, shortString: "GMT-03:00"},
    {label: "(GMT-03:00) Chile Standard Time – Santiago", minutesDiff: -180, shortString: "GMT-03:00"},
    {label: "(GMT-03:00) Falkland Islands Time – Stanley", minutesDiff: -180, shortString: "GMT-03:00"},
    {label: "(GMT-03:00) French Guiana Time – Cayenne", minutesDiff: -180, shortString: "GMT-03:00"},
    {label: "(GMT-03:00) Paraguay Standard Time – Asunción", minutesDiff: -180, shortString: "GMT-03:00"},
    {label: "(GMT-03:00) Rothera Time – Rothera", minutesDiff: -180, shortString: "GMT-03:00"},
    {label: "(GMT-03:00) Suriname Time – Paramaribo", minutesDiff: -180, shortString: "GMT-03:00"},
    {label: "(GMT-03:00) Uruguay Standard Time – Montevideo", minutesDiff: -180, shortString: "GMT-03:00"},
    {label: "(GMT-03:00) West Greenland Time – Nuuk", minutesDiff: -180, shortString: "GMT-03:00"},
    {label: "(GMT-02:00) Brasilia Time – Noronha", minutesDiff: -120, shortString: "GMT-02:00"},
    {label: "(GMT-02:00) South Georgia Time – South Georgia", minutesDiff: -120, shortString: "GMT-02:00"},
    {label: "(GMT-01:00) Azores Time – Azores", minutesDiff: -60, shortString: "GMT-01:00"},
    {label: "(GMT+00:00) Coordinated Universal Time", minutesDiff: 0, shortString: "GMT+00:00"},
    {label: "(GMT+00:00) Greenwich Mean Time – Abidjan", minutesDiff: 0, shortString: "GMT+00:00"},
    {label: "(GMT+00:00) Greenwich Mean Time – Bissau", minutesDiff: 0, shortString: "GMT+00:00"},
    {label: "(GMT+00:00) Greenwich Mean Time – Danmarkshavn", minutesDiff: 0, shortString: "GMT+00:00"},
    {label: "(GMT+00:00) Greenwich Mean Time – Monrovia", minutesDiff: 0, shortString: "GMT+00:00"},
    {label: "(GMT+00:00) Greenwich Mean Time – São Tomé", minutesDiff: 0, shortString: "GMT+00:00"},
    {label: "(GMT+01:00) Central European Standard Time – Algiers", minutesDiff: 60, shortString: "GMT+01:00"},
    {label: "(GMT+01:00) Central European Standard Time – Tunis", minutesDiff: 60, shortString: "GMT+01:00"},
    {label: "(GMT+01:00) United Kingdom Time – London", minutesDiff: 60, shortString: "GMT+01:00"},
    {label: "(GMT+01:00) Ireland Time – Dublin", minutesDiff: 60, shortString: "GMT+01:00"},
    {label: "(GMT+01:00) Morocco Time – Casablanca", minutesDiff: 60, shortString: "GMT+01:00"},
    {label: "(GMT+01:00) West Africa Standard Time – Lagos", minutesDiff: 60, shortString: "GMT+01:00"},
    {label: "(GMT+01:00) West Africa Standard Time – Ndjamena", minutesDiff: 60, shortString: "GMT+01:00"},
    {label: "(GMT+01:00) Western European Time – Canary", minutesDiff: 60, shortString: "GMT+01:00"},
    {label: "(GMT+01:00) Western European Time – Faroe", minutesDiff: 60, shortString: "GMT+01:00"},
    {label: "(GMT+01:00) Western European Time – Lisbon", minutesDiff: 60, shortString: "GMT+01:00"},
    {label: "(GMT+01:00) Western European Time – Madeira", minutesDiff: 60, shortString: "GMT+01:00"},
    {label: "(GMT+01:00) Western Sahara Time – El Aaiun", minutesDiff: 60, shortString: "GMT+01:00"},
    {label: "(GMT+02:00) Central Africa Time – Maputo", minutesDiff: 120, shortString: "GMT+02:00"},
    {label: "(GMT+02:00) Central Africa Time – Juba", minutesDiff: 120, shortString: "GMT+02:00"},
    {label: "(GMT+02:00) Central Africa Time – Khartoum", minutesDiff: 120, shortString: "GMT+02:00"},
    {label: "(GMT+02:00) Central Africa Time – Windhoek", minutesDiff: 120, shortString: "GMT+02:00"},
    {label: "(GMT+02:00) Central European Time – Brussels", minutesDiff: 120, shortString: "GMT+02:00"},
    {label: "(GMT+02:00) Central European Time – Andorra", minutesDiff: 120, shortString: "GMT+02:00"},
    {label: "(GMT+02:00) Central European Time – Belgrade", minutesDiff: 120, shortString: "GMT+02:00"},
    {label: "(GMT+02:00) Central European Time – Berlin", minutesDiff: 120, shortString: "GMT+02:00"},
    {label: "(GMT+02:00) Central European Time – Prague", minutesDiff: 120, shortString: "GMT+02:00"},
    {label: "(GMT+02:00) Central European Time – Budapest", minutesDiff: 120, shortString: "GMT+02:00"},
    {label: "(GMT+02:00) Central European Time – Zurich", minutesDiff: 120, shortString: "GMT+02:00"},
    {label: "(GMT+02:00) Central European Time – Ceuta", minutesDiff: 120, shortString: "GMT+02:00"},
    {label: "(GMT+02:00) Central European Time – Gibraltar", minutesDiff: 120, shortString: "GMT+02:00"},
    {label: "(GMT+02:00) Central European Time – Madrid", minutesDiff: 120, shortString: "GMT+02:00"},
    {label: "(GMT+02:00) Central European Time – Malta", minutesDiff: 120, shortString: "GMT+02:00"},
    {label: "(GMT+02:00) Central European Time – Paris", minutesDiff: 120, shortString: "GMT+02:00"},
    {label: "(GMT+02:00) Central European Time – Rome", minutesDiff: 120, shortString: "GMT+02:00"},
    {label: "(GMT+02:00) Central European Time – Tirane", minutesDiff: 120, shortString: "GMT+02:00"},
    {label: "(GMT+02:00) Central European Time – Vienna", minutesDiff: 120, shortString: "GMT+02:00"},
    {label: "(GMT+02:00) Central European Time – Warsaw", minutesDiff: 120, shortString: "GMT+02:00"},
    {label: "(GMT+02:00) Eastern European Standard Time – Kaliningrad", minutesDiff: 120, shortString: "GMT+02:00"},
    {label: "(GMT+02:00) Eastern European Standard Time – Tripoli", minutesDiff: 120, shortString: "GMT+02:00"},
    {label: "(GMT+02:00) South Africa Standard Time – Johannesburg", minutesDiff: 120, shortString: "GMT+02:00"},
    {label: "(GMT+02:00) Troll Time – Troll", minutesDiff: 120, shortString: "GMT+02:00"},
    {label: "(GMT+03:00) Arabian Standard Time – Riyadh", minutesDiff: 180, shortString: "GMT+03:00"},
    {label: "(GMT+03:00) Arabian Standard Time – Baghdad", minutesDiff: 180, shortString: "GMT+03:00"},
    {label: "(GMT+03:00) Arabian Standard Time – Qatar", minutesDiff: 180, shortString: "GMT+03:00"},
    {label: "(GMT+03:00) East Africa Time – Nairobi", minutesDiff: 180, shortString: "GMT+03:00"},
    {label: "(GMT+03:00) Eastern European Time – Athens", minutesDiff: 180, shortString: "GMT+03:00"},
    {label: "(GMT+03:00) Eastern European Time – Beirut", minutesDiff: 180, shortString: "GMT+03:00"},
    {label: "(GMT+03:00) Eastern European Time – Bucharest", minutesDiff: 180, shortString: "GMT+03:00"},
    {label: "(GMT+03:00) Eastern European Time – Cairo", minutesDiff: 180, shortString: "GMT+03:00"},
    {label: "(GMT+03:00) Eastern European Time – Chisinau", minutesDiff: 180, shortString: "GMT+03:00"},
    {label: "(GMT+03:00) Eastern European Time – Gaza", minutesDiff: 180, shortString: "GMT+03:00"},
    {label: "(GMT+03:00) Eastern European Time – Hebron", minutesDiff: 180, shortString: "GMT+03:00"},
    {label: "(GMT+03:00) Eastern European Time – Helsinki", minutesDiff: 180, shortString: "GMT+03:00"},
    {label: "(GMT+03:00) Eastern European Time – Kyiv", minutesDiff: 180, shortString: "GMT+03:00"},
    {label: "(GMT+03:00) Eastern European Time – Nicosia", minutesDiff: 180, shortString: "GMT+03:00"},
    {label: "(GMT+03:00) Eastern European Time – Riga", minutesDiff: 180, shortString: "GMT+03:00"},
    {label: "(GMT+03:00) Eastern European Time – Sofia", minutesDiff: 180, shortString: "GMT+03:00"},
    {label: "(GMT+03:00) Eastern European Time – Tallinn", minutesDiff: 180, shortString: "GMT+03:00"},
    {label: "(GMT+03:00) Eastern European Time – Vilnius", minutesDiff: 180, shortString: "GMT+03:00"},
    {label: "(GMT+03:00) Famagusta Time – Famagusta", minutesDiff: 180, shortString: "GMT+03:00"},
    {label: "(GMT+03:00) Israel Time – Jerusalem", minutesDiff: 180, shortString: "GMT+03:00"},
    {label: "(GMT+03:00) Jordan Time – Amman", minutesDiff: 180, shortString: "GMT+03:00"},
    {label: "(GMT+03:00) Kirov Time – Kirov", minutesDiff: 180, shortString: "GMT+03:00"},
    {label: "(GMT+03:00) Moscow Standard Time – Minsk", minutesDiff: 180, shortString: "GMT+03:00"},
    {label: "(GMT+03:00) Moscow Standard Time – Moscow", minutesDiff: 180, shortString: "GMT+03:00"},
    {label: "(GMT+03:00) Moscow Standard Time – Simferopol", minutesDiff: 180, shortString: "GMT+03:00"},
    {label: "(GMT+03:00) Syria Time – Damascus", minutesDiff: 180, shortString: "GMT+03:00"},
    {label: "(GMT+03:00) Türkiye Time – Istanbul", minutesDiff: 180, shortString: "GMT+03:00"},
    {label: "(GMT+03:00) Volgograd Standard Time – Volgograd", minutesDiff: 180, shortString: "GMT+03:00"},
    {label: "(GMT+03:30) Iran Standard Time – Tehran", minutesDiff: 210, shortString: "GMT+03:30"},
    {label: "(GMT+04:00) Armenia Standard Time – Yerevan", minutesDiff: 240, shortString: "GMT+04:00"},
    {label: "(GMT+04:00) Astrakhan Time – Astrakhan", minutesDiff: 240, shortString: "GMT+04:00"},
    {label: "(GMT+04:00) Azerbaijan Standard Time – Baku", minutesDiff: 240, shortString: "GMT+04:00"},
    {label: "(GMT+04:00) Georgia Standard Time – Tbilisi", minutesDiff: 240, shortString: "GMT+04:00"},
    {label: "(GMT+04:00) Gulf Standard Time – Dubai", minutesDiff: 240, shortString: "GMT+04:00"},
    {label: "(GMT+04:00) Mauritius Standard Time – Mauritius", minutesDiff: 240, shortString: "GMT+04:00"},
    {label: "(GMT+04:00) Samara Standard Time – Samara", minutesDiff: 240, shortString: "GMT+04:00"},
    {label: "(GMT+04:00) Saratov Time – Saratov", minutesDiff: 240, shortString: "GMT+04:00"},
    {label: "(GMT+04:00) Ulyanovsk Time – Ulyanovsk", minutesDiff: 240, shortString: "GMT+04:00"},
    {label: "(GMT+04:30) Afghanistan Time – Kabul", minutesDiff: 270, shortString: "GMT+04:30"},
    {label: "(GMT+05:00) Maldives Time – Maldives", minutesDiff: 300, shortString: "GMT+05:00"},
    {label: "(GMT+05:00) Mawson Time – Mawson", minutesDiff: 300, shortString: "GMT+05:00"},
    {label: "(GMT+05:00) Pakistan Standard Time – Karachi", minutesDiff: 300, shortString: "GMT+05:00"},
    {label: "(GMT+05:00) Tajikistan Time – Dushanbe", minutesDiff: 300, shortString: "GMT+05:00"},
    {label: "(GMT+05:00) Turkmenistan Standard Time – Ashgabat", minutesDiff: 300, shortString: "GMT+05:00"},
    {label: "(GMT+05:00) Uzbekistan Standard Time – Samarkand", minutesDiff: 300, shortString: "GMT+05:00"},
    {label: "(GMT+05:00) Uzbekistan Standard Time – Tashkent", minutesDiff: 300, shortString: "GMT+05:00"},
    {label: "(GMT+05:00) Vostok Time – Vostok", minutesDiff: 300, shortString: "GMT+05:00"},
    {label: "(GMT+05:00) West Kazakhstan Time – Almaty", minutesDiff: 300, shortString: "GMT+05:00"},
    {label: "(GMT+05:00) West Kazakhstan Time – Aqtau", minutesDiff: 300, shortString: "GMT+05:00"},
    {label: "(GMT+05:00) West Kazakhstan Time – Aqtobe", minutesDiff: 300, shortString: "GMT+05:00"},
    {label: "(GMT+05:00) West Kazakhstan Time – Atyrau", minutesDiff: 300, shortString: "GMT+05:00"},
    {label: "(GMT+05:00) West Kazakhstan Time – Kostanay", minutesDiff: 300, shortString: "GMT+05:00"},
    {label: "(GMT+05:00) West Kazakhstan Time – Oral", minutesDiff: 300, shortString: "GMT+05:00"},
    {label: "(GMT+05:00) West Kazakhstan Time – Qyzylorda", minutesDiff: 300, shortString: "GMT+05:00"},
    {label: "(GMT+05:00) Yekaterinburg Standard Time – Yekaterinburg", minutesDiff: 300, shortString: "GMT+05:00"},
    {label: "(GMT+05:30) India Standard Time – Colombo", minutesDiff: 330, shortString: "GMT+05:30"},
    {label: "(GMT+05:30) India Standard Time – Kolkata", minutesDiff: 330, shortString: "GMT+05:30"},
    {label: "(GMT+05:45) Nepal Time – Kathmandu", minutesDiff: 345, shortString: "GMT+05:45"},
    {label: "(GMT+06:00) Bangladesh Standard Time – Dhaka", minutesDiff: 360, shortString: "GMT+06:00"},
    {label: "(GMT+06:00) Bhutan Time – Thimphu", minutesDiff: 360, shortString: "GMT+06:00"},
    {label: "(GMT+06:00) Indian Ocean Time – Chagos", minutesDiff: 360, shortString: "GMT+06:00"},
    {label: "(GMT+06:00) Kyrgyzstan Time – Bishkek", minutesDiff: 360, shortString: "GMT+06:00"},
    {label: "(GMT+06:00) Omsk Standard Time – Omsk", minutesDiff: 360, shortString: "GMT+06:00"},
    {label: "(GMT+06:00) Urumqi Time – Urumqi", minutesDiff: 360, shortString: "GMT+06:00"},
    {label: "(GMT+06:30) Myanmar Time – Yangon", minutesDiff: 390, shortString: "GMT+06:30"},
    {label: "(GMT+07:00) Barnaul Time – Barnaul", minutesDiff: 420, shortString: "GMT+07:00"},
    {label: "(GMT+07:00) Indochina Time – Bangkok", minutesDiff: 420, shortString: "GMT+07:00"},
    {label: "(GMT+07:00) Davis Time – Davis", minutesDiff: 420, shortString: "GMT+07:00"},
    {label: "(GMT+07:00) Hovd Standard Time – Hovd", minutesDiff: 420, shortString: "GMT+07:00"},
    {label: "(GMT+07:00) Indochina Time – Ho Chi Minh City", minutesDiff: 420, shortString: "GMT+07:00"},
    {label: "(GMT+07:00) Krasnoyarsk Standard Time – Krasnoyarsk", minutesDiff: 420, shortString: "GMT+07:00"},
    {label: "(GMT+07:00) Krasnoyarsk Standard Time – Novokuznetsk", minutesDiff: 420, shortString: "GMT+07:00"},
    {label: "(GMT+07:00) Novosibirsk Standard Time – Novosibirsk", minutesDiff: 420, shortString: "GMT+07:00"},
    {label: "(GMT+07:00) Tomsk Time – Tomsk", minutesDiff: 420, shortString: "GMT+07:00"},
    {label: "(GMT+07:00) Western Indonesia Time – Jakarta", minutesDiff: 420, shortString: "GMT+07:00"},
    {label: "(GMT+07:00) Western Indonesia Time – Pontianak", minutesDiff: 420, shortString: "GMT+07:00"},
    {label: "(GMT+08:00) Australian Western Standard Time – Casey", minutesDiff: 480, shortString: "GMT+08:00"},
    {label: "(GMT+08:00) Australian Western Standard Time – Perth", minutesDiff: 480, shortString: "GMT+08:00"},
    {label: "(GMT+08:00) Malaysia Time – Kuching", minutesDiff: 480, shortString: "GMT+08:00"},
    {label: "(GMT+08:00) Central Indonesia Time – Makassar", minutesDiff: 480, shortString: "GMT+08:00"},
    {label: "(GMT+08:00) China Standard Time – Macao", minutesDiff: 480, shortString: "GMT+08:00"},
    {label: "(GMT+08:00) China Standard Time – Shanghai", minutesDiff: 480, shortString: "GMT+08:00"},
    {label: "(GMT+08:00) Hong Kong Standard Time – Hong Kong", minutesDiff: 480, shortString: "GMT+08:00"},
    {label: "(GMT+08:00) Irkutsk Standard Time – Irkutsk", minutesDiff: 480, shortString: "GMT+08:00"},
    {label: "(GMT+08:00) Singapore Standard Time – Singapore", minutesDiff: 480, shortString: "GMT+08:00"},
    {label: "(GMT+08:00) Philippine Standard Time – Manila", minutesDiff: 480, shortString: "GMT+08:00"},
    {label: "(GMT+08:00) Taipei Standard Time – Taipei", minutesDiff: 480, shortString: "GMT+08:00"},
    {label: "(GMT+08:00) Ulaanbaatar Standard Time – Choibalsan", minutesDiff: 480, shortString: "GMT+08:00"},
    {label: "(GMT+08:00) Ulaanbaatar Standard Time – Ulaanbaatar", minutesDiff: 480, shortString: "GMT+08:00"},
    {label: "(GMT+08:45) Australian Central Western Standard Time – Eucla", minutesDiff: 525, shortString: "GMT+08:45"},
    {label: "(GMT+09:00) East Timor Time – Dili", minutesDiff: 540, shortString: "GMT+09:00"},
    {label: "(GMT+09:00) Eastern Indonesia Time – Jayapura", minutesDiff: 540, shortString: "GMT+09:00"},
    {label: "(GMT+09:00) Japan Standard Time – Tokyo", minutesDiff: 540, shortString: "GMT+09:00"},
    {label: "(GMT+09:00) Korean Standard Time – Pyongyang", minutesDiff: 540, shortString: "GMT+09:00"},
    {label: "(GMT+09:00) Korean Standard Time – Seoul", minutesDiff: 540, shortString: "GMT+09:00"},
    {label: "(GMT+09:00) Palau Time – Palau", minutesDiff: 540, shortString: "GMT+09:00"},
    {label: "(GMT+09:00) Yakutsk Standard Time – Chita", minutesDiff: 540, shortString: "GMT+09:00"},
    {label: "(GMT+09:00) Yakutsk Standard Time – Khandyga", minutesDiff: 540, shortString: "GMT+09:00"},
    {label: "(GMT+09:00) Yakutsk Standard Time – Yakutsk", minutesDiff: 540, shortString: "GMT+09:00"},
    {label: "(GMT+09:30) Australian Central Standard Time – Darwin", minutesDiff: 570, shortString: "GMT+09:30"},
    {label: "(GMT+09:30) Central Australia Time – Adelaide", minutesDiff: 570, shortString: "GMT+09:30"},
    {label: "(GMT+09:30) Central Australia Time – Broken Hill", minutesDiff: 570, shortString: "GMT+09:30"},
    {label: "(GMT+10:00) Australian Eastern Standard Time – Brisbane", minutesDiff: 600, shortString: "GMT+10:00"},
    {label: "(GMT+10:00) Australian Eastern Standard Time – Lindeman", minutesDiff: 600, shortString: "GMT+10:00"},
    {label: "(GMT+10:00) Chamorro Standard Time – Guam", minutesDiff: 600, shortString: "GMT+10:00"},
    {label: "(GMT+10:00) Papua New Guinea Time – Port Moresby", minutesDiff: 600, shortString: "GMT+10:00"},
    {label: "(GMT+10:00) Eastern Australia Time – Hobart", minutesDiff: 600, shortString: "GMT+10:00"},
    {label: "(GMT+10:00) Eastern Australia Time – Macquarie", minutesDiff: 600, shortString: "GMT+10:00"},
    {label: "(GMT+10:00) Eastern Australia Time – Melbourne", minutesDiff: 600, shortString: "GMT+10:00"},
    {label: "(GMT+10:00) Eastern Australia Time – Sydney", minutesDiff: 600, shortString: "GMT+10:00"},
    {label: "(GMT+10:00) Vladivostok Standard Time – Ust-Nera", minutesDiff: 600, shortString: "GMT+10:00"},
    {label: "(GMT+10:00) Vladivostok Standard Time – Vladivostok", minutesDiff: 600, shortString: "GMT+10:00"},
    {label: "(GMT+10:30) Lord Howe Time – Lord Howe", minutesDiff: 630, shortString: "GMT+10:30"},
    {label: "(GMT+11:00) Bougainville Time – Bougainville", minutesDiff: 660, shortString: "GMT+11:00"},
    {label: "(GMT+11:00) Kosrae Time – Kosrae", minutesDiff: 660, shortString: "GMT+11:00"},
    {label: "(GMT+11:00) Magadan Standard Time – Magadan", minutesDiff: 660, shortString: "GMT+11:00"},
    {label: "(GMT+11:00) New Caledonia Standard Time – Noumea", minutesDiff: 660, shortString: "GMT+11:00"},
    {label: "(GMT+11:00) Norfolk Island Time – Norfolk", minutesDiff: 660, shortString: "GMT+11:00"},
    {label: "(GMT+11:00) Solomon Islands Time – Guadalcanal", minutesDiff: 660, shortString: "GMT+11:00"},
    {label: "(GMT+11:00) Sakhalin Standard Time – Sakhalin", minutesDiff: 660, shortString: "GMT+11:00"},
    {label: "(GMT+11:00) Srednekolymsk Time – Srednekolymsk", minutesDiff: 660, shortString: "GMT+11:00"},
    {label: "(GMT+11:00) Vanuatu Standard Time – Efate", minutesDiff: 660, shortString: "GMT+11:00"},
    {label: "(GMT+12:00) Anadyr Standard Time – Anadyr", minutesDiff: 720, shortString: "GMT+12:00"},
    {label: "(GMT+12:00) Fiji Standard Time – Fiji", minutesDiff: 720, shortString: "GMT+12:00"},
    {label: "(GMT+12:00) Gilbert Islands Time – Tarawa", minutesDiff: 720, shortString: "GMT+12:00"},
    {label: "(GMT+12:00) Marshall Islands Time – Kwajalein", minutesDiff: 720, shortString: "GMT+12:00"},
    {label: "(GMT+12:00) Nauru Time – Nauru", minutesDiff: 720, shortString: "GMT+12:00"},
    {label: "(GMT+12:00) New Zealand Time – Auckland", minutesDiff: 720, shortString: "GMT+12:00"},
    {label: "(GMT+12:00) Petropavlovsk-Kamchatski Standard Time – Kamchatka", minutesDiff: 720, shortString: "GMT+12:00"},
    {label: "(GMT+12:45) Chatham Time – Chatham", minutesDiff: 765, shortString: "GMT+12:45"},
    {label: "(GMT+13:00) Apia Standard Time – Apia", minutesDiff: 780, shortString: "GMT+13:00"},
    {label: "(GMT+13:00) Phoenix Islands Time – Enderbury", minutesDiff: 780, shortString: "GMT+13:00"},
    {label: "(GMT+13:00) Tokelau Time – Fakaofo", minutesDiff: 780, shortString: "GMT+13:00"},
    {label: "(GMT+13:00) Tonga Standard Time – Tongatapu", minutesDiff: 780, shortString: "GMT+13:00"},
    {label: "(GMT+14:00) Line Islands Time – Kiritimati", minutesDiff: 840, shortString: "GMT+14:00"}
  ];


  //per test ne prendo solo 5
  // timezones = timezones.slice(0, 4);
  // timezones.push({label: "(GMT+02:00) Central European Time – Rome", minutesDiff: 120, shortString: "GMT+02:00"});
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

    const refTimezoneStr = "(GMT+02:00) Central European Time – Rome";
    const refTimezone = timezones.find(tz => tz.label === refTimezoneStr);
    //va convertita la data nel fuso orario corrente
    const localDates = data.map(d => new Date(d.DateInfo.getTime() + timezone.minutesDiff * 60000 + refTimezone.minutesDiff * 60000));
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
/* ====================================================================
   CONFIG
   ==================================================================== */
const DAILY_LIMIT   = 100;   // stop here
const DAILY_WARNING = 90;    // warn here
const TRACKER_NAME  = 'Org/Biz Outreach Tracker';   // <— your sheet tab
/* ==================================================================== */

/* ========== 0.  API-KEY SETTER ===================================== */
function setGooglePlacesApiKey() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt(
    'Set Google Places API Key',
    'Enter a valid key that starts with "AIza…":',
    ui.ButtonSet.OK_CANCEL
  );
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const key = res.getResponseText().trim();
  if (!key) return;
  PropertiesService.getScriptProperties().setProperty('PLACES_API_KEY', key);
  ui.alert('Key saved! You can now fetch businesses.');
}

/* prompt if missing, else return key */
function ensureApiKey_() {
  let key = PropertiesService.getScriptProperties().getProperty('PLACES_API_KEY');
  if (key) return key;
  setGooglePlacesApiKey();
  return PropertiesService.getScriptProperties().getProperty('PLACES_API_KEY');
}

/* ========== 1.  DAILY QUOTA COUNTER ================================ */
function checkAndCountRequests_(needed) {
  const props = PropertiesService.getScriptProperties();
  const today = Utilities.formatDate(new Date(), 'GMT', 'yyyy-MM-dd');
  let meta;
  try { meta = JSON.parse(props.getProperty('PLACES_USAGE') || '{}'); }
  catch(e){ meta = {}; }

  if (meta.date !== today) meta = { date: today, count: 0 };

  if (meta.count + needed > DAILY_LIMIT) {
    SpreadsheetApp.getUi().alert(
      `Daily Google Places limit (${DAILY_LIMIT}) reached.\nTry again tomorrow.`
    );
    return false;
  }
  if (meta.count < DAILY_WARNING && meta.count + needed >= DAILY_WARNING) {
    SpreadsheetApp.getUi().alert(
      `Heads-up: you’ve used ${DAILY_WARNING}+ of ${DAILY_LIMIT} requests today.`
    );
  }
  meta.count += needed;
  props.setProperty('PLACES_USAGE', JSON.stringify(meta));
  return true;
}

/* ========== 2.  SIMPLE EMAIL HELPERS ================================ */
function scrapeEmailFromSite_(url) {
  if (!url) return '';
  try {
    const html = UrlFetchApp.fetch(url, {muteHttpExceptions:true, followRedirects:true, timeout:8000})
                 .getContentText();
    const m = html.match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/ig);
    if (m && m.length) return m[0].toLowerCase();
  } catch(e){}
  return '';
}

function hunterEmail_(domain){
  const key = PropertiesService.getScriptProperties().getProperty('HUNTER_API_KEY');
  if (!key || !domain) return '';
  try{
    const r = UrlFetchApp.fetch(
      `https://api.hunter.io/v2/domain-search?domain=${domain}&api_key=${key}&limit=1`
    );
    const d = JSON.parse(r.getContentText());
    return (d.data.emails && d.data.emails[0] && d.data.emails[0].value) || '';
  }catch(e){ return ''; }
}

/* ========== 3.  CREATE TRACKER SHEET =============================== */
function createSponsorshipOutreachTracker() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSheetByName(TRACKER_NAME)) {
    SpreadsheetApp.getUi().alert('Tracker sheet already exists.'); return;
  }
  const sh = ss.insertSheet(TRACKER_NAME, 0);

  /* dashboard rows 1-7 */
  const dash = [
    ['ORG/BIZ OUTREACH DASHBOARD','','','','','','',''],
    ['Total Leads','=COUNTA(A9:A)-1',
     'Confirmed','=COUNTIF(Q9:Q,"Confirmed")',
     'Interested','=COUNTIF(Q9:Q,"Interested")',
     'Declined','=COUNTIF(Q9:Q,"Declined")'],
    ['Prod. Donations','=COUNTIF(R9:R,"Product Donation")',
     'Gift Cards','=COUNTIF(R9:R,"Gift Card")',
     'Cash','=COUNTIF(R9:R,"Cash Sponsorship")',
     'Volunteers','=COUNTIF(R9:R,"Volunteer Support")'],
    ['','','','','','','',''],
    ['','','','','','','',''],
    ['','','','','','','',''],
    ['','','','','','','','']
  ];
  sh.getRange(1,1,7,8).setValues(dash);
  sh.getRange(1,1,1,8).setFontSize(16).setFontWeight('bold').setBackground('#b7dde8');
  sh.getRange(1,1,7,8).setBackground('#f4f9fb');

  /* headers row 8 */
  const headers = [
    'Business Name','Primary Category','Full Address','City','State','ZIP',
    'Phone','Website','Google Rating','# Reviews','Business Status','Maps Link',
    'Place ID','Contact Person','Email','Date Contacted','Response Status',
    'Type of Support','Follow-Up Date','Notes'
  ];
  sh.getRange(8,1,1,headers.length).setValues([headers])
    .setFontWeight('bold').setBackground('#dce6f2');

  sh.setFrozenRows(8);
  sh.getRange(9,1,1000,headers.length)
    .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY)
    .getHeaderColumnRange().setBackground('#dce6f2');

  const widths=[220,150,260,120,60,80,150,220,90,90,140,220,300,150,200,120,140,160,120,240];
  widths.forEach((w,i)=>sh.setColumnWidth(i+1,w));
  sh.hideColumns(13);  // place ID

  const dateRule = SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).build();
  sh.getRange('P9:P1000').setDataValidation(dateRule);
  sh.getRange('S9:S1000').setDataValidation(dateRule);

  const statusRule=SpreadsheetApp.newDataValidation()
        .requireValueInList(['Not Contacted','Contacted','Interested','Confirmed','Declined'],true)
        .setAllowInvalid(false).build();
  sh.getRange('Q9:Q1000').setDataValidation(statusRule);

  const supportRule=SpreadsheetApp.newDataValidation()
        .requireValueInList(['Product Donation','Gift Card','Cash Sponsorship','Volunteer Support','Other'],true)
        .setAllowInvalid(false).build();
  sh.getRange('R9:R1000').setDataValidation(supportRule);

  const cfR=sh.getRange('Q9:Q1000'), colours={'Confirmed':'#b6d7a8','Interested':'#a4c2f4','Contacted':'#ffe599','Declined':'#f4cccc','Not Contacted':'#d9d9d9'};
  sh.setConditionalFormatRules(Object.keys(colours).map(v=>
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(v).setBackground(colours[v]).setRanges([cfR]).build()
  ));
  sh.getRange(8,1,1000,headers.length).createFilter();
}

/* ========== 4.  SIDEBAR =========================================== */
function showSearchSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('SearchForm')
               .setTitle('Add Nearby Businesses');
  SpreadsheetApp.getUi().showSidebar(html);
}

/* ========== 5.  MAIN FETCH & APPEND =============================== */
function addNearbyBusinesses(formData) {
  const { term, zip, maxResults } = formData;
  const apiKey = ensureApiKey_();
  if (!apiKey) { SpreadsheetApp.getUi().alert('No API key provided.'); return ''; }

  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TRACKER_NAME);
  if (!sh) throw new Error('Tracker sheet missing.');

  const duplicates = new Set(sh.getRange('M9:M').getValues().flat().filter(String));
  let url = `https://maps.googleapis.com/maps/api/place/textsearch/json?query=${encodeURIComponent(term+' in '+zip)}&key=${apiKey}`;
  let fetched=0, rows=[], page=0;

  while (url && fetched<maxResults) {
    if (!checkAndCountRequests_(1)) return 'Quota exceeded.';
    const resp=JSON.parse(UrlFetchApp.fetch(url).getContentText());
    if (resp.status!=='OK') break;

    for (const p of resp.results) {
      if (fetched>=maxResults) break;
      if (duplicates.has(p.place_id)) continue;

      if (!checkAndCountRequests_(1)) return 'Quota exceeded.';
      const det=JSON.parse(UrlFetchApp.fetch(
          `https://maps.googleapis.com/maps/api/place/details/json?place_id=${p.place_id}&fields=name,formatted_phone_number,website&key=${apiKey}`
      ).getContentText()).result||{};

      const addr=p.formatted_address||'', parts=addr.split(',');
      let city='',state='',zipParsed='';
      if(parts.length>=3){
        city=parts[parts.length-2].trim();
        const s=parts[parts.length-1].trim().split(' ');
        state=s[0]||''; zipParsed=s[1]||'';
      }

      let email='';
      if(det.website){
        email=scrapeEmailFromSite_(det.website);
        if(!email){
          const domain=det.website.replace(/https?:\/\/(www\.)?/i,'').split(/[\/?#]/)[0];
          email=hunterEmail_(domain);
        }
      }

      rows.push([
        p.name||'', (p.types&&p.types[0])||'', addr, city, state, zipParsed,
        det.formatted_phone_number||'',
        det.website||`https://google.com/maps/place/?q=place_id:${p.place_id}`,
        p.rating||'', p.user_ratings_total||'', p.business_status||'',
        `https://google.com/maps/place/?q=place_id:${p.place_id}`,
        p.place_id,'',email,'','','','',''
      ]);
      fetched++;
      Utilities.sleep(200);
    }

    url=(resp.next_page_token&&fetched<maxResults)
       ? `https://maps.googleapis.com/maps/api/place/textsearch/json?pagetoken=${resp.next_page_token}&key=${apiKey}`
       : null;
    if(url) Utilities.sleep(2000);
    page++; if(page===3) url=null;
  }

  if(!rows.length) return 'No new businesses (all duplicates).';
  sh.getRange(sh.getLastRow()+1,1,rows.length,rows[0].length).setValues(rows);
  return `Added ${rows.length} new businesses!`;
}


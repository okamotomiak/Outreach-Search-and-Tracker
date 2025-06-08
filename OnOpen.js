/**
 * Adds a “Sponsor Tools” menu on spreadsheet open.
 * – 🔧  Create Outreach Tracker  →  builds the sheet (row 8 header, dashboard)
 * – 📍  Add Nearby Businesses (ZIP)  →  opens the sidebar search form
 * – 🔑  Set Google Places API Key    →  lets you enter / replace the key
 * – ⭐  Update Lead Scores           →  (optional) run the scoring function
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Sponsor Tools')
      .addItem('🔧 Create Outreach Tracker', 'createSponsorshipOutreachTracker')
      .addItem('📍 Add Nearby Businesses (ZIP)', 'showSearchSidebar')
      .addSeparator()
      .addItem('🔑 Set Google Places API Key', 'setGooglePlacesApiKey')
    .addToUi();
}

/**
 * Sort bulk mail into Gmail labels: "Bulk Newsletters", "Bulk Ads", "Bulk Other"
 * Creates the labels if they don't exist, then labels matching messages.
 *
 * HOW TO USE:
 *   1. Go to https://script.google.com  →  New Project
 *   2. Paste this entire file into the editor (replace any existing code)
 *   3. Click the function dropdown (top toolbar) and select:
 *        - "dryRun"   to count without labeling
 *        - "sortBulkMail" to actually label and archive messages
 *   4. Click ▶ Run  (first run will ask for Gmail permissions — approve it)
 *   5. View → Execution log to see results
 *
 * NOTE: Messages are labeled AND archived (removed from Inbox).
 *       They are NOT deleted — find them under their new label anytime.
 *       To change this behavior, set ARCHIVE_AFTER_LABEL = false below.
 */

const ARCHIVE_AFTER_LABEL = true;  // set false to label but keep in Inbox

// -------------------------------------------------------
// Categorized sender lists (same as your Outlook script)
// -------------------------------------------------------
const categories = {
  'Bulk Newsletters': [
    'newsletters@emails.discoveryplus.com',
    'packtdatapro1@substack.com',
    'membership@outside.plus',
    'email@washingtonpost.com',
    'newsletter@mail.coinbase.com',
    'support@email.masterclass.com',
    'info@perforce.com',
    'justin.reock@perforce.com',
    'newsletter@email.lifecare-news.com',
    'membership@gaiagps.com',
    'john@predictabledesigns.com',
  ],
  'Bulk Ads': [
    // Retail / shopping
    'hft@em.harborfreight.com',
    'OmahaSteaks@mail.omahasteaks.com',
    'GameStop@emails.gamestop.com',
    'samsung@innovations.samsungusa.com',
    'sales@impactguns.com',
    'zenhelp@omcaviar.com',
    'fandango@movies.fandango.com',
    'contact@swingbyswing.com',
    'course@golffacility.com',
    // Financial promos
    'hello@capitaloneshopping.com',
    'capitalone@offer.capitalone.com',
    'capitalone@message.capitalone.com',
    'donotreply@cardmessage.capitalone.com',
    'Fidelity.Investments@mail.fidelity.com',
    'reply@email-firsthorizon.com',
    'help@selflender.com',
    'eDelivery@etradefrommorganstanley.com',
    // Drip / nurture campaigns
    'clientcare@vu.com',
    'education@vu.com',
    'team@vu.com',
    // Software promos
    'norton@secure.norton.com',
    'Autodesk@autodeskcommunications.com',
    'quicken@mail.quicken.com',
    'VSPVisionCareVCM@e.vsp.com',
  ],
  'Bulk Other': [
    // Recruiting spam / job boards
    'dailyjobalert@postjobfree.com',
    'donotreply@upwork.com',
    'enrollment@wgu.edu',
    'team@email.rocketlawyer.com',
    // Misc bulk
    'sales@bitraser.com',
    'noreply@github.com',
    'noreply@medium.com',
    'noreply@quora.com',
    'support@nordvpn.com',
  ],
};

// -------------------------------------------------------
// Main: Sort bulk mail (labels + optional archive)
// -------------------------------------------------------
function sortBulkMail() {
  processBulkMail_(false);
}

// -------------------------------------------------------
// Dry run: Count matches without changing anything
// -------------------------------------------------------
function dryRun() {
  processBulkMail_(true);
}

// -------------------------------------------------------
// Core logic
// -------------------------------------------------------
function processBulkMail_(isDryRun) {
  Logger.log(isDryRun ? '=== DRY RUN MODE ===' : '=== SORTING BULK MAIL ===');
  Logger.log('');

  let grandTotal = 0;
  const results = {};

  for (const [catName, senders] of Object.entries(categories)) {
    Logger.log('--- ' + catName + ' ---');
    const label = isDryRun ? null : getOrCreateLabel_(catName);
    let catTotal = 0;

    for (const sender of senders) {
      // Gmail search: from:sender
      const query = 'from:' + sender;
      const threads = findAllThreads_(query);

      if (threads.length > 0) {
        const msgCount = threads.reduce((sum, t) => sum + t.getMessageCount(), 0);
        Logger.log('  ' + padLeft_(msgCount, 5) + '  ' + sender + '  (' + threads.length + ' threads)');
        catTotal += msgCount;

        if (!isDryRun) {
          // Process in batches to avoid timeout
          for (let i = 0; i < threads.length; i += 100) {
            const batch = threads.slice(i, i + 100);
            // Add label
            label.addToThreads(batch);
            // Archive (remove from Inbox) if configured
            if (ARCHIVE_AFTER_LABEL) {
              GmailApp.moveThreadsToArchive(batch);
            }
          }
        }
      }
    }

    results[catName] = catTotal;
    grandTotal += catTotal;
    Logger.log('  Total for ' + catName + ': ' + catTotal + ' messages');
    Logger.log('');
  }

  Logger.log('=======================================');
  Logger.log('Grand total messages: ' + grandTotal);
  Logger.log('=======================================');
  Logger.log('');

  for (const [cat, count] of Object.entries(results)) {
    Logger.log(cat + ': ' + count);
  }

  if (isDryRun) {
    Logger.log('');
    Logger.log('DRY RUN — no messages were labeled or moved.');
    Logger.log('Select "sortBulkMail" and run again to apply.');
  } else {
    Logger.log('');
    Logger.log('Done! Check your Gmail labels:');
    Logger.log('  • Bulk Newsletters');
    Logger.log('  • Bulk Ads');
    Logger.log('  • Bulk Other');
    if (ARCHIVE_AFTER_LABEL) {
      Logger.log('Messages have been archived (removed from Inbox).');
    }
  }
}

// -------------------------------------------------------
// Helper: Get all matching threads (handles 500-thread limit)
// -------------------------------------------------------
function findAllThreads_(query) {
  const allThreads = [];
  let start = 0;
  const pageSize = 500;

  while (true) {
    const batch = GmailApp.search(query, start, pageSize);
    if (batch.length === 0) break;
    allThreads.push(...batch);
    if (batch.length < pageSize) break;
    start += pageSize;
  }

  return allThreads;
}

// -------------------------------------------------------
// Helper: Get or create a Gmail label
// -------------------------------------------------------
function getOrCreateLabel_(labelName) {
  let label = GmailApp.getUserLabelByName(labelName);
  if (label) {
    Logger.log('  Label "' + labelName + '" exists');
    return label;
  }
  label = GmailApp.createLabel(labelName);
  Logger.log('  Created label "' + labelName + '"');
  return label;
}

// -------------------------------------------------------
// Helper: Left-pad a number
// -------------------------------------------------------
function padLeft_(num, width) {
  const s = String(num);
  return s.length >= width ? s : ' '.repeat(width - s.length) + s;
}

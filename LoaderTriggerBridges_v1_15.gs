/**
 * LOTTEON Apps Script local trigger bridges v1.15
 *
 * Install this file beside Loader_v1_14_main_full.gs in the bound Apps Script
 * project. The daily 06:10 scheduler creates a trigger whose handler must
 * exist in the local project; remote eval functions are not persisted between
 * trigger executions.
 */
const LOTTEON_TRIGGER_BRIDGE_VERSION = 'v1.15';

/** Local entrypoint used by startDailyFilterCountsSchedule(). */
function runDailyFilterCountsStart() {
  return runFilterCountLightweightFunction_('runDailyFilterCountsStart');
}

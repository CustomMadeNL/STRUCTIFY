/**
 * STRUCTIFY — Full module builder (P-003 … P-007)
 * ==================================================================
 * Generates every module for the five expansion packs as proper
 * multi-tab Google Sheets in the STRUCTIFY_WIP folder, matching the
 * house layout: an INDEX tab + one tab per section, each with its own
 * header row, banding, column widths and working formulas.
 *
 * HOW TO RUN
 *   1. script.google.com  ->  New project
 *   2. Paste this whole file, Save.
 *   3. Run  buildAllModules   (approve the Drive permission once).
 *      -> ~23 Sheets appear in STRUCTIFY_WIP.
 *   Tip: run buildOne('M-021') to (re)build a single module while testing.
 *
 * Re-running creates duplicates — delete old copies first, or call
 * deleteExistingModuleCopies() once before a rebuild.
 *
 * Modules (from STRUCTIFY_MODULE_REGISTRY):
 *   P-003 All You Need : M-001 M-002 M-003 M-004 M-005 M-045 M-046 M-037 M-038
 *   P-004 Money Control: M-021 M-022 M-023 M-024 M-025
 *   P-005 Kids         : M-029 M-030 M-031 M-032
 *   P-006 Family       : M-026 M-027 M-028 (+M-029 +M-031 shared)
 *   P-007 SHE-O        : M-048 M-047 (+M-046 +M-021 +M-037 shared)
 */

var WIP_FOLDER_ID = '1YOVFHy--GQ_81U49b3oY8Fx_qwtXNCib'; // STRUCTIFY_WIP
var FONT = 'Montserrat';                                  // brand font
var C = { black:'#000000', white:'#ffffff' };
var MONEY = '€ #,##0.00';

// ======================================================== run entrypoints ==
function buildAllModules() { MODULES.forEach(function (m) { buildModule_(m); }); }
function buildOne(id)      { buildModule_(MODULES.filter(function(m){return m.id===id;})[0]); }

function deleteExistingModuleCopies() {
  var folder = DriveApp.getFolderById(WIP_FOLDER_ID);
  MODULES.forEach(function (m) {
    var it = folder.getFilesByName(m.id + '_' + m.short);
    while (it.hasNext()) it.next().setTrashed(true);
  });
}

// ============================================================= build core ==
function buildModule_(m) {
  var title = m.id + '_' + m.short;
  var ss = SpreadsheetApp.create(title);
  var file = DriveApp.getFileById(ss.getId());
  DriveApp.getFolderById(WIP_FOLDER_ID).addFile(file);
  DriveApp.getRootFolder().removeFile(file);

  // OVERVIEW = default sheet, plain SECTION | CONTENT table (header on row 1)
  fillTab_(ss.getSheets()[0].setName('OVERVIEW'), {
    headers:['SECTION','CONTENT'],
    rows:[['MODULE', m.name],['CODE', m.id + ' · ' + m.pack],
          ['PURPOSE', m.purpose],['HOW TO USE', m.howto],
          ['RULE', m.rule],['STATUS','SAMPLE (WIP)']],
    widths:[18,100]
  });

  // one tab per section — tab names UPPERCASE_WITH_UNDERSCORES
  m.tabs.forEach(function (t) {
    fillTab_(ss.insertSheet(t.title.replace(/ /g,'_')), t);
  });
}

// Fills a given sheet: header on ROW 1 (bold white on black, frozen),
// data from row 2, Montserrat 10, no banding/borders — matches M-000.
function fillTab_(sh, t) {
  var nCols = t.headers.length;
  sh.getRange(1,1,1,nCols).setValues([t.headers])
    .setBackground(C.black).setFontColor(C.white).setFontWeight('bold')
    .setFontFamily(FONT).setFontSize(10);
  sh.setFrozenRows(1);

  var first = 2;
  var rows = t.rows || [];
  var last = first + rows.length - 1;
  if (rows.length) {
    sh.getRange(first,1,rows.length,nCols).setValues(pad_(rows,nCols))
      .setFontFamily(FONT).setFontSize(10).setVerticalAlignment('top').setWrap(true);
  }

  // per-row formulas: [{col, tpl}]  tpl uses {r}
  (t.rowFormulas||[]).forEach(function (rf) {
    for (var r=first;r<=last;r++) sh.getRange(r,rf.col).setFormula(sub_(rf.tpl,{r:r}));
  });

  // summary rows after a blank line; strings with {first}/{last} become formulas
  if (t.summaryRows && t.summaryRows.length) {
    var sr = last + 2;
    t.summaryRows.forEach(function (row, i) {
      for (var c=0;c<row.length;c++) {
        var v = row[c], cell = sh.getRange(sr+i, c+1);
        if (typeof v==='string' && v.charAt(0)==='=') cell.setFormula(sub_(v,{first:first,last:last}));
        else cell.setValue(v);
        cell.setFontFamily(FONT).setFontSize(10);
        if (c===0) cell.setFontWeight('bold');
      }
    });
  }

  if (t.widths) t.widths.forEach(function (w,i){ sh.setColumnWidth(i+1,w); });
  if (t.moneyCols) t.moneyCols.forEach(function (c){ sh.getRange(first,c,Math.max(rows.length,1),1).setNumberFormat(MONEY); });
  return sh;
}

function pad_(rows,n){ return rows.map(function(r){ var a=r.slice(); while(a.length<n)a.push(''); return a; }); }
function sub_(tpl,map){ return tpl.replace(/\{(\w+)\}/g,function(_,k){ return map[k]; }); }

// ================================================================ configs ==
var MODULES = [
// --------------------------------------------------------------- P-003 -----
{ id:'M-001', short:'LIFE_DASHBOARD', name:'LIFE DASHBOARD', pack:'P-003 All You Need System',
  purpose:'See where you stand in every area of your life and choose one focus at a time.',
  howto:'Rate each area 1-10, choose ONE focus, set a weekly priority, review weekly.',
  rule:'One area first. Improve one thing, then the next.',
  tabs:[
    { title:'DASHBOARD', subtitle:'Rate each area, choose one focus, log the next action.',
      headers:['LIFE AREA','SATISFACTION (1-10)','CURRENT FOCUS','STATUS','NEXT ACTION'],
      widths:[24,18,34,16,40],
      rows:[
        ['Health & Energy',7,'Sleep before 23:00','On track','Set a wind-down alarm'],
        ['Career & Work',6,'Finish STRUCTIFY launch','Focus','Publish P-001 on Payhip'],
        ['Money',5,'Build a buffer','Needs work','Set up the budget tracker'],
        ['Relationships',8,'Weekly date night','On track','Plan Friday evening'],
        ['Personal Growth',6,'Read 20 min/day','Focus',"Pick this month's book"],
        ['Fun & Recreation',5,'Protect weekends','Needs work','Block Saturday morning'],
        ['Home & Environment',7,'Declutter one room','On track','Start with the desk'],
        ['Purpose & Meaning',6,'Clarify the 1-year goal','Focus','Write it in the Goal System']],
      summaryRows:[
        ['AVERAGE SATISFACTION','=ROUND(AVERAGE(B{first}:B{last}),1)'],
        ['AREAS NEEDING WORK (<6)','=COUNTIF(B{first}:B{last},"<6")']] },
    { title:'WEEKLY FOCUS', subtitle:'Your single priority this week.',
      headers:['ITEM','YOUR ANSWER'], widths:[24,60],
      rows:[["This week's #1 focus",'Publish P-001 (Your AI Agent Pack)'],
        ['Why it matters','First live product = first revenue'],
        ['Areas in focus','Career & Work, Money'],
        ['One thing to say no to','New product ideas until P-001 is live'],
        ['Weekly review day','Friday']] }]},

{ id:'M-002', short:'GOAL_SYSTEM', name:'GOAL SYSTEM', pack:'P-003 All You Need System',
  purpose:'Turn vague ambitions into clear goals with milestones and a status you can track.',
  howto:'Write each goal, link it to a life area, set a deadline and break it into milestones.',
  rule:'Three active goals maximum. Focus beats a long wish list.',
  tabs:[
    { title:'GOALS', subtitle:'One row per goal. Keep it to a few active goals.',
      headers:['GOAL','LIFE AREA','WHY IT MATTERS','DEADLINE','NEXT MILESTONE','STATUS'],
      widths:[30,18,32,14,30,14],
      rows:[
        ['Launch STRUCTIFY','Career','Create income and proof','2026-09-30','Publish P-001','In progress'],
        ['3-month buffer','Money','Peace of mind','2026-12-31','Save first €500','In progress'],
        ['Run 5k','Health','Energy and focus','2026-11-01','Run 2k non-stop','Not started']] }]},

{ id:'M-003', short:'WEEKLY_PLANNING', name:'WEEKLY PLANNING', pack:'P-003 All You Need System',
  purpose:'Plan your week around what matters most — one clear top priority per day.',
  howto:'Set your top 3 for the week, then one Top Priority per day with tasks and appointments.',
  rule:'One top priority per day. If everything is a priority, nothing is.',
  tabs:[
    { title:'WEEK', subtitle:'One clear top priority per day.',
      headers:['DAY','TOP PRIORITY','TASKS','APPOINTMENTS','NOTES'], widths:[12,28,34,26,26],
      rows:[
        ['Monday','Finalize P-001 delivery PDF','Paste links; export PDF','Deep work 9-11',''],
        ['Tuesday','','','',''],['Wednesday','','','',''],['Thursday','','','',''],
        ['Friday','Publish P-001 on Payhip','Upload PDF; set price; test order','Weekly review 16:00','Launch day'],
        ['Saturday','','','',''],['Sunday','','','','']] },
    { title:'BRAIN DUMP', subtitle:'Empty your head here, then sort into the week.',
      headers:['EVERYTHING ON YOUR MIND','CATEGORY','DO WHEN'], widths:[60,16,16],
      rows:[['Reply to supplier email','Work','Today'],['Book dentist','Life','This week'],
        ['Idea: bundle P-001 + P-002','Business','Later'],['','',''],['','',''],['','','']] }]},

{ id:'M-004', short:'DAILY_PLANNING', name:'DAILY PLANNING', pack:'P-003 All You Need System',
  purpose:'Structure a single day around your top 3 tasks and a simple time plan.',
  howto:'Set your top 3, then block your day. Tick tasks as done.',
  rule:'Top 3 first. Do the most important thing before noon.',
  tabs:[
    { title:'TODAY', subtitle:'Top 3, then the timeline.',
      headers:['TIME','FOCUS / TASK','DONE'], widths:[14,50,10],
      rows:[['Top 3','1. Publish P-001  2. Draft mockups  3. Reply supplier',''],
        ['08:00','Deep work — delivery PDF',''],['10:00','Email + admin',''],
        ['13:00','Mockups in Canva',''],['15:00','Publish + test order',''],
        ['17:00','Shut down + plan tomorrow','']] }]},

{ id:'M-005', short:'HABIT_TRACKER', name:'HABIT TRACKER', pack:'P-003 All You Need System',
  purpose:'Build habits that stick by tracking a few key ones each week — not twenty at once.',
  howto:'List your habits, set a weekly goal, mark each day you do it. DONE and % update automatically.',
  rule:'Start with 2-3 habits. Consistency beats ambition.',
  tabs:[
    { title:'TRACKER', subtitle:'Mark each day. Totals are automatic.',
      headers:['HABIT','GOAL /WK','MON','TUE','WED','THU','FRI','SAT','SUN','DONE','% OF GOAL'],
      widths:[26,10,6,6,6,6,6,6,6,8,10],
      rows:[
        ['Move 20 min',5,'x','x','','x','x','','x'],
        ['Read 20 min',7,'x','x','x','','x','x','x'],
        ['No phone before bed',7,'','x','x','x','','x','x'],
        ['Plan tomorrow',5,'x','x','x','x','x','',''],
        ['Drink 2L water',7,'x','x','x','x','x','x','x']],
      rowFormulas:[{col:10,tpl:'=COUNTA(C{r}:I{r})'},{col:11,tpl:'=ROUND(COUNTA(C{r}:I{r})/B{r}*100,0)'}] }]},

{ id:'M-045', short:'PERSONAL_CRM', name:'PERSONAL CRM', pack:'P-003 All You Need System',
  purpose:'Stay in touch with the people who matter instead of letting months slip by.',
  howto:'List people, note the last time you connected and when to reach out next.',
  rule:'Relationships are a system too. Small, regular touches beat rare big ones.',
  tabs:[
    { title:'PEOPLE', subtitle:'Who matters, and when to reach out next.',
      headers:['NAME','RELATIONSHIP','LAST CONTACT','NEXT TOUCH','NOTES'], widths:[22,20,16,16,40],
      rows:[['Anna','Mentor','2026-07-20','2026-08-20','Send launch update'],
        ['Team CustomMade','Work','2026-08-14','2026-08-18','Weekly sync'],
        ['Mom','Family','2026-08-10','2026-08-17','Call Sunday']] }]},

{ id:'M-046', short:'ENERGY_MANAGEMENT', name:'ENERGY MANAGEMENT', pack:'P-003 All You Need System',
  purpose:'Plan your work around your energy instead of forcing focus at the wrong times.',
  howto:'Map your typical energy through the day and match tasks to each block.',
  rule:'Protect your peak hours for your most important work.',
  tabs:[
    { title:'ENERGY MAP', subtitle:'Match the work to the energy.',
      headers:['TIME BLOCK','ENERGY (1-5)','BEST FOR','WHAT DRAINS ME'], widths:[18,14,34,30],
      rows:[['Early morning',5,'Deep work, writing','Checking email first'],
        ['Late morning',4,'Meetings, calls','Back-to-back with no break'],
        ['Early afternoon',2,'Admin, light tasks','Big decisions'],
        ['Late afternoon',3,'Creative, planning','Doomscrolling']] }]},

{ id:'M-037', short:'WEEKLY_REVIEW', name:'WEEKLY REVIEW', pack:'P-003 All You Need System',
  purpose:'Close each week with a short review so you learn and reset instead of drifting.',
  howto:'Answer the prompts every week (10 minutes). Roll unfinished priorities forward.',
  rule:'A weekly review is the habit that makes every other system work.',
  tabs:[
    { title:'REVIEW', subtitle:'10 minutes, every week.',
      headers:['PROMPT','YOUR ANSWER'], widths:[36,64],
      rows:[['Wins this week',''],['What worked',''],['What to drop',''],
        ['Unfinished — roll over',''],['Top 3 for next week',''],['One improvement','']] }]},

{ id:'M-038', short:'MONTHLY_REVIEW', name:'MONTHLY REVIEW', pack:'P-003 All You Need System',
  purpose:'Zoom out monthly to see progress on goals and adjust course.',
  howto:'Answer the prompts at month end and set the theme for next month.',
  rule:'Weeks are for doing, months are for steering.',
  tabs:[
    { title:'REVIEW', subtitle:'Zoom out and steer.',
      headers:['PROMPT','YOUR ANSWER'], widths:[36,64],
      rows:[['Biggest win',''],['Progress on goals',''],['What slipped',''],
        ['Numbers that matter (sales, savings)',''],['Theme for next month',''],['One thing to change','']] }]},

// --------------------------------------------------------------- P-004 -----
{ id:'M-021', short:'BUDGET_TRACKER', name:'BUDGET TRACKER', pack:'P-004 Money Control System',
  purpose:'Plan your monthly budget and see planned vs actual per category at a glance.',
  howto:'Set a planned amount per category, log the actual, and the difference is automatic.',
  rule:'Give every euro a job before the month starts.',
  tabs:[
    { title:'BUDGET', subtitle:'Planned vs actual, per category.',
      headers:['CATEGORY','PLANNED','ACTUAL','DIFFERENCE'], widths:[28,16,16,16], moneyCols:[2,3,4],
      rows:[['Housing',1200,1200,''],['Groceries',400,435,''],['Transport',150,120,''],
        ['Subscriptions',60,72,''],['Fun & eating out',200,240,''],['Savings',300,300,'']],
      rowFormulas:[{col:4,tpl:'=B{r}-C{r}'}],
      summaryRows:[['TOTAL','=SUM(B{first}:B{last})','=SUM(C{first}:C{last})','=SUM(D{first}:D{last})']] }]},

{ id:'M-022', short:'EXPENSE_TRACKER', name:'EXPENSE TRACKER', pack:'P-004 Money Control System',
  purpose:'Log spending as it happens and see where the money actually goes.',
  howto:'Add a row per expense. Totals update at the bottom.',
  rule:'You can only manage what you can see. Log it the same day.',
  tabs:[
    { title:'EXPENSES', subtitle:'One row per expense.',
      headers:['DATE','DESCRIPTION','CATEGORY','AMOUNT'], widths:[14,40,20,16], moneyCols:[4],
      rows:[['2026-08-01','Groceries — Albert Heijn','Groceries',68.40],
        ['2026-08-02','Train to client','Transport',24.10],
        ['2026-08-03','Netflix','Subscriptions',12.99],
        ['2026-08-04','Dinner out','Fun & eating out',52.00]],
      summaryRows:[['','','TOTAL','=SUM(D{first}:D{last})']] }]},

{ id:'M-023', short:'CASHFLOW_DASHBOARD', name:'CASHFLOW DASHBOARD', pack:'P-004 Money Control System',
  purpose:'See money in versus money out each month and whether you are net positive.',
  howto:'Enter income and expenses per month; net is calculated automatically.',
  rule:'Spend less than you earn — then automate the gap into savings.',
  tabs:[
    { title:'CASHFLOW', subtitle:'Money in vs money out.',
      headers:['MONTH','INCOME','EXPENSES','NET'], widths:[16,16,16,16], moneyCols:[2,3,4],
      rows:[['June',3200,2900,''],['July',3400,3100,''],['August',3300,3050,'']],
      rowFormulas:[{col:4,tpl:'=B{r}-C{r}'}],
      summaryRows:[['TOTAL','=SUM(B{first}:B{last})','=SUM(C{first}:C{last})','=SUM(D{first}:D{last})']] }]},

{ id:'M-024', short:'SAVINGS_SYSTEM', name:'SAVINGS SYSTEM', pack:'P-004 Money Control System',
  purpose:'Set savings goals and watch how close you are to each one.',
  howto:'Enter a target and what you have saved; remaining and % are automatic.',
  rule:'Pay yourself first — automate savings on payday.',
  tabs:[
    { title:'SAVINGS GOALS', subtitle:'Targets and progress.',
      headers:['GOAL','TARGET','SAVED','REMAINING','% DONE'], widths:[28,16,16,16,10], moneyCols:[2,3,4],
      rows:[['Emergency buffer',3000,1200,'',''],['Holiday',1500,600,'',''],['New laptop',1400,350,'','']],
      rowFormulas:[{col:4,tpl:'=B{r}-C{r}'},{col:5,tpl:'=ROUND(C{r}/B{r}*100,0)'}] }]},

{ id:'M-025', short:'SUBSCRIPTION_TRACKER', name:'SUBSCRIPTION TRACKER', pack:'P-004 Money Control System',
  purpose:'Catch the recurring costs quietly draining your account.',
  howto:'List each subscription and its monthly cost; the monthly total is automatic.',
  rule:'Review quarterly and cancel anything you did not use last month.',
  tabs:[
    { title:'SUBSCRIPTIONS', subtitle:'Everything recurring, in one place.',
      headers:['NAME','MONTHLY COST','RENEWAL DAY','USED LAST MONTH?','KEEP?'], widths:[24,16,14,20,10], moneyCols:[2],
      rows:[['Netflix',12.99,'12','Yes','Keep'],['Spotify',10.99,'03','Yes','Keep'],
        ['Adobe',24.19,'20','Rarely','Review'],['Gym',39.00,'01','No','Cancel']],
      summaryRows:[['TOTAL / MONTH','=SUM(B{first}:B{last})']] }]},

// --------------------------------------------------------------- P-005 -----
{ id:'M-029', short:'KIDS_ROUTINE_SYSTEM', name:'KIDS ROUTINE SYSTEM', pack:'P-005 STRUCTIFY Kids',
  purpose:'Give kids simple, visual morning and evening routines they can follow themselves.',
  howto:'Fill in the steps for each routine. Let the child tick them off.',
  rule:'Keep it short and visual. 5 steps beat 15.',
  tabs:[
    { title:'ROUTINES', subtitle:'Morning and evening, step by step.',
      headers:['TIME','MORNING ROUTINE','EVENING ROUTINE'], widths:[12,40,40],
      rows:[['Step 1','Get dressed','Pyjamas on'],['Step 2','Breakfast','Brush teeth'],
        ['Step 3','Brush teeth','Pack bag for tomorrow'],['Step 4','Pack school bag','Read 10 min'],
        ['Step 5','Shoes & go','Lights out']] }]},

{ id:'M-030', short:'HOMEWORK_SYSTEM', name:'HOMEWORK SYSTEM', pack:'P-005 STRUCTIFY Kids',
  purpose:'Keep homework visible so nothing is forgotten and evenings are calmer.',
  howto:'Add each task with its due date and tick when done.',
  rule:'Write it down the day it is given, not the night before.',
  tabs:[
    { title:'HOMEWORK', subtitle:'Every task, its due date, and status.',
      headers:['SUBJECT','TASK','DUE','EST. TIME','DONE'], widths:[18,36,14,12,10],
      rows:[['Maths','Exercises p.24','2026-08-18','30 min',''],
        ['Reading','Chapter 3','2026-08-19','20 min',''],
        ['History','Project research','2026-08-22','45 min','']] }]},

{ id:'M-031', short:'SCREEN_TIME_SYSTEM', name:'SCREEN TIME SYSTEM', pack:'P-005 STRUCTIFY Kids',
  purpose:'Turn screen-time into clear, agreed rules so it stops being a daily fight.',
  howto:'Agree the rules together and log daily allowed vs earned time.',
  rule:'Agree the rules once, calmly — then just follow the sheet.',
  tabs:[
    { title:'AGREEMENT', subtitle:'The rules, agreed together.',
      headers:['RULE','AGREED'], widths:[50,30],
      rows:[['Screens after homework and chores','Yes'],['No screens at the table','Yes'],
        ['Devices charge outside the bedroom','Yes'],['Extra time is earned, not default','Yes']] },
    { title:'DAILY LOG', subtitle:'Allowed vs earned.',
      headers:['DAY','ALLOWED (min)','EARNED (min)','NOTES'], widths:[12,14,14,34],
      rows:[['Mon',60,30,''],['Tue',60,0,''],['Wed',60,45,''],['Thu',60,30,''],
        ['Fri',90,60,''],['Sat',120,0,''],['Sun',120,0,'']] }]},

{ id:'M-032', short:'KIDS_GOAL_TRACKER', name:'KIDS GOAL TRACKER', pack:'P-005 STRUCTIFY Kids',
  purpose:'Help kids set small goals and see their progress toward a reward.',
  howto:'Set a goal and a reward; colour in the progress as they go.',
  rule:'Small goals, quick rewards. Momentum matters more than size.',
  tabs:[
    { title:'GOALS', subtitle:'Goals and rewards.',
      headers:['GOAL','REWARD','PROGRESS (of 5)','DONE?'], widths:[34,24,16,10],
      rows:[['Read 5 books','Movie night','3',''],['Tidy room 5 days','Ice cream','2',''],
        ['Practice piano 5x','New sticker set','4','']] }]},

// --------------------------------------------------------------- P-006 -----
{ id:'M-026', short:'FAMILY_CALENDAR', name:'FAMILY CALENDAR', pack:'P-006 Family Operating System',
  purpose:'One shared weekly view of where everyone needs to be.',
  howto:'Fill in each day by part of day. Share the copy with the family.',
  rule:'One shared calendar beats five private ones.',
  tabs:[
    { title:'WEEK', subtitle:'Everyone, at a glance.',
      headers:['DAY','MORNING','AFTERNOON','EVENING'], widths:[12,30,30,30],
      rows:[['Monday','School','Football (Sam)','Family dinner'],['Tuesday','School','',''],
        ['Wednesday','School','Swimming (Mia)',''],['Thursday','School','',''],
        ['Friday','School','','Movie night'],['Saturday','Groceries','Birthday party',''],
        ['Sunday','','Grandparents','Prep week']] }]},

{ id:'M-027', short:'MEAL_PLANNER', name:'MEAL PLANNER', pack:'P-006 Family Operating System',
  purpose:'Plan the week of meals and build the grocery list from it.',
  howto:'Plan dinners for the week, then add what you need to the grocery list.',
  rule:'Plan seven dinners on Sunday and shop once.',
  tabs:[
    { title:'MEALS', subtitle:'Dinners for the week.',
      headers:['DAY','DINNER','NOTES'], widths:[12,36,30],
      rows:[['Monday','Pasta bolognese',''],['Tuesday','Stir-fry & rice',''],
        ['Wednesday','Soup & bread','Quick night'],['Thursday','Tacos',''],
        ['Friday','Pizza','Homemade'],['Saturday','BBQ',''],['Sunday','Roast chicken','']] },
    { title:'GROCERIES', subtitle:'Built from the meal plan.',
      headers:['ITEM','QTY','GOT IT?'], widths:[36,10,10],
      rows:[['Minced meat','500g',''],['Pasta','1 pack',''],['Vegetables (mixed)','',''],
        ['Tortillas','1 pack',''],['Chicken','1',''],['','','']] }]},

{ id:'M-028', short:'CHORE_SYSTEM', name:'CHORE SYSTEM', pack:'P-006 Family Operating System',
  purpose:'Share household chores fairly and make it clear who does what, when.',
  howto:'Assign each chore to a person and a day; tick when done.',
  rule:'Everyone contributes. Make it visible and rotate.',
  tabs:[
    { title:'CHORES', subtitle:'Who does what, when.',
      headers:['CHORE','WHO','FREQUENCY','DAY','DONE'], widths:[28,16,16,14,10],
      rows:[['Take out bins','Sam','Weekly','Tuesday',''],['Vacuum','Mia','Weekly','Saturday',''],
        ['Kitchen clean','Rotating','Daily','—',''],['Laundry','Parent','2x/week','Wed & Sun',''],
        ['Set the table','Kids','Daily','—','']] }]},

// --------------------------------------------------------------- P-007 -----
{ id:'M-048', short:'SHE_O_DASHBOARD', name:'SHE-O DASHBOARD', pack:'P-007 SHE-O Operating System',
  purpose:'One calm overview of life and business — status and next action per area.',
  howto:'Set this week for each area, mark the status and the single next action.',
  rule:'Run the week from one screen. Life and business live together here.',
  tabs:[
    { title:'DASHBOARD', subtitle:'Life + business, one view.',
      headers:['AREA','THIS WEEK','STATUS','NEXT ACTION'], widths:[20,34,16,34],
      rows:[['Business','Land 2 discovery calls','On track','Send 5 outreach messages'],
        ['Money','Invoice this month','Focus','Send invoice #0042'],
        ['Energy','Protect deep-work mornings','Needs work','Block 8-10 daily'],
        ['Boundaries','No work after 18:30','Focus','Turn off notifications'],
        ['Self','Two workouts','On track','Book Tue & Thu']] }]},

{ id:'M-047', short:'BOUNDARY_SYSTEM', name:'BOUNDARY SYSTEM', pack:'P-007 SHE-O Operating System',
  purpose:'Define your boundaries and the exact words to hold them.',
  howto:'For each area, write your boundary and a ready-to-use phrase.',
  rule:'A boundary you can say out loud is a boundary you can keep.',
  tabs:[
    { title:'BOUNDARIES', subtitle:'Your lines, and the words to hold them.',
      headers:['AREA','MY BOUNDARY','WHAT I SAY','STATUS'], widths:[20,34,40,14],
      rows:[['Work hours','No calls after 18:30',"\"I'll pick this up first thing tomorrow.\"",'Holding'],
        ['Scope creep','No unpaid extras',"\"Happy to — let's add it as a small extra.\"",'Working on it'],
        ['Personal time','Sundays offline',"\"I'm offline Sundays, back Monday.\"",'Holding']] }]}
];

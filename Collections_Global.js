const PROJECT_NAME = "Paybox - DPU Monitoring v2";

/**
 * Collections Helper Functions
 * This file contains utility functions for managing collection operations
 */

// === Constants === //

/**
 * Mapping of day indices to day abbreviations
 * @type {Object.<number, string>}
 */
const dayMapping = {
  0: 'Sun.',
  1: 'M.',
  2: 'T.',
  3: 'W.',
  4: 'Th.',
  5: 'F.',
  6: 'Sat.'
};

/**
 * Mapping of day abbreviations to day indices
 * @type {Object.<string, number>}
 */
const dayIndex = {
  'Sun.': 0,
  'M.': 1,
  'T.': 2,
  'W.': 3,
  'Th.': 4,
  'F.': 5,
  'Sat.': 6
};

/**
 * Amount thresholds for collection by day
 * @type {Object.<string, number>}
 */
const amountThresholds = {
  'M.': 300000,
  'T.': 310000,
  'W.': 310000,
  'Th.': 300000,
  'F.': 290000,
  'Sat.': 290000,
  'Sun.': 290000
};

// const amountThresholds = {
//   'M.': 250000,
//   'T.': 250000,
//   'W.': 250000,
//   'Th.': 250000,
//   'F.': 250000,
//   'Sat.': 250000,
//   'Sun.': 250000
// };

/**
 * Payday ranges for collection
 * @type {Array.<{start: number, end: number}>}
 */
const paydayRanges = [
  { start: 15, end: 16 },
  { start: 30, end: 31 }
];

/**
 * Due date cutoffs for collection
 * @type {Array.<{start: number, end: number}>}
 */
const dueDateCutoffs = [
  { start: 5, end: 6 },
  { start: 20, end: 21 }
];

/**
 * Amount threshold for payday collections
 * @type {number}
 */
const paydayAmount = 290000;
// const paydayAmount = 250000;

/**
 * Amount threshold for due date collections
 * @type {number}
 */
const dueDateCutoffsAmount = 290000;
// const dueDateCutoffsAmount = 250000;

/**
 * Email signature for all emails
 * @type {string}
 */
const emailSignature = `<div><div dir="ltr" class="gmail_signature"><div dir="ltr"><span><div dir="ltr" style="margin-left:0pt" align="left">Best Regards,</div><div dir="ltr" style="margin-left:0pt" align="left"><br><table style="border:none;border-collapse:collapse"><colgroup><col width="44"><col width="249"><col width="100"><col width="7"></colgroup><tbody><tr style="height:66.75pt"><td rowspan="3" style="vertical-align:top;background-color:rgb(249,249,249);overflow:hidden"><p dir="ltr" style="line-height:1.2;text-align:justify;margin-top:0pt;margin-bottom:0pt"><span style="font-size:11pt;font-family:Arial,sans-serif;color:rgb(0,0,0);vertical-align:baseline"><span style="border:none;display:inline-block;overflow:hidden;width:47px;height:217px"><img src="https://lh3.googleusercontent.com/JNGaTKS3JTfQlHCnlwFXo14_knhu4v_WhlZCWOIFJfRPuUKjMMHWuj82yUQ0uUOxv9XNk1Nooae__kDJ1wS0st_Xe3SZvDdl3dkVSpX24SCtgfIt7ZfeTfIR8S93ndcLMdQSgm9Xyq1rykUOGv1sLo0" width="47" height="235.00000000000003" style="margin-left:0px;margin-top:0px" crossorigin="" class="CToWUd" data-bit="iit"></span></span></p></td><td style="vertical-align:middle;background-color:rgb(249,249,249);overflow:hidden"><p dir="ltr" style="line-height:1.38;margin-top:0pt;margin-bottom:0pt"><span style="font-size:12pt;font-family:Arial,sans-serif;background-color:transparent;font-weight:700;vertical-align:baseline">Support</span></p><p dir="ltr" style="line-height:1.38;margin-top:0pt;margin-bottom:0pt"><span style="font-size:12pt;font-family:Arial,sans-serif;background-color:transparent;font-weight:700;vertical-align:baseline">Paybox Operations</span></p><br></td><td colspan="2" rowspan="3" style="vertical-align:top;background-color:rgb(249,249,249);overflow:hidden"><p dir="ltr" style="line-height:1.2;margin-right:0.75pt;text-align:justify;margin-top:0pt;margin-bottom:0pt"><span style="font-size:11pt;font-family:Arial,sans-serif;color:rgb(0,0,0);vertical-align:baseline"><span style="border:none;display:inline-block;overflow:hidden;width:118px;height:240px"><img src="https://lh6.googleusercontent.com/3QMqLONmIp2CUDF9DoGaWmEhai3RSB6fjgqFxwhxwcQoX58wlvnAVNMscDRgfOK-xv4S2bllMTzrKQSuvgqAi68syHzvqbNJziibdwTfx7A1pSWelqdkffPtJ9n6WC3JJEEcSqNYXrBthmb8cxIz5Dg" width="118" height="240" style="margin-left:0px;margin-top:0px" crossorigin="" class="CToWUd" data-bit="iit"></span></span></p></td></tr><tr style="height:84pt"><td style="vertical-align:top;background-color:rgb(249,249,249);overflow:hidden"><p dir="ltr" style="line-height:1.8;margin-top:0pt;margin-bottom:0pt"><span style="font-size:8pt;font-family:Arial,sans-serif;color:rgb(0,0,0);vertical-align:baseline"><span style="border:none;display:inline-block;overflow:hidden;width:14px;height:10px"><img src="https://lh4.googleusercontent.com/DM0kOfMz47NSJfROHj6mLS0ypJih5nHezY-SaBODOfd_oXMKxDagoXJG1WmGYaCgt0g_PLa4KQY1Btkuih7F2409F3-gjDxV8UGVeL_6bKF4l3Aze7QG33MalyKV0NmslPNz5aK3Fp8a8LX_8abfJoo" width="14" height="10" style="margin-left:0px;margin-top:0px" crossorigin="" class="CToWUd" data-bit="iit"></span></span><span style="font-size:8pt;font-family:Arial,sans-serif;color:rgb(0,0,0);vertical-align:baseline"> </span><a href="mailto:support@paybox.ph" target="_blank"><span style="font-size:10pt;font-family:Arial,sans-serif;color:rgb(17,85,204);vertical-align:baseline">support@paybox.ph</span></a></p><p dir="ltr" style="line-height:1.8;margin-top:0pt;margin-bottom:0pt"><span style="font-size:8pt;font-family:Arial,sans-serif;color:rgb(0,0,0);vertical-align:baseline"><span style="border:none;display:inline-block;overflow:hidden;width:15px;height:16px"><img src="https://lh6.googleusercontent.com/RFRWIimtGNr-ErnlmxYSpbrwRjPjBvZXbY5BdP3LD2ykCtMVGMYodZoIc-B7xHWXI3wYHcAr8FxK6d3L4hk12mbH-dTEZ1pU6pugBvzZeqvu2uLo_4BPb_zlAlT8ve3P2GD0CifMeZ_dX_qdWhPns5g" width="15" height="16" style="margin-left:0px;margin-top:0px" crossorigin="" class="CToWUd" data-bit="iit"></span></span><span style="font-size:8pt;font-family:Arial,sans-serif;color:rgb(0,0,0);vertical-align:baseline"> </span><span style="font-size:10pt;font-family:Arial,sans-serif;color:rgb(0,0,0);vertical-align:baseline">+62 (2) 8835-9697</span></p><p dir="ltr" style="line-height:1.8;margin-top:0pt;margin-bottom:0pt"><span style="font-size:8pt;font-family:Arial,sans-serif;color:rgb(0,0,0);vertical-align:baseline"><span style="border:none;display:inline-block;overflow:hidden;width:15px;height:13px"><img src="https://lh4.googleusercontent.com/VuSTEwKlvSLg3DHUziWOwLBgmj1ctniwh9f7tsnECrzxxdMy1CRKnKgJaCi2m8xIRfrsdtsihCEvexpSHnDykgsRZ1WeMuxwHKQpbS-VUfBwoM2wC3oZU9i7B8vgAWf4JZPzq03GND-SOi9bD0RIIJM" width="15" height="13" style="margin-left:0px;margin-top:0px" crossorigin="" class="CToWUd" data-bit="iit"></span></span><span style="font-size:8pt;font-family:Arial,sans-serif;color:rgb(0,0,0);vertical-align:baseline"> </span><span style="font-size:10pt;font-family:Arial,sans-serif;color:rgb(0,0,0);vertical-align:baseline">+62 (2) 8805-1066 (loc) 115&nbsp;</span></p><p dir="ltr" style="line-height:1.8;margin-top:0pt;margin-bottom:0pt"><span style="font-size:8pt;font-family:Arial,sans-serif;color:rgb(0,0,0);vertical-align:baseline"><span style="border:none;display:inline-block;overflow:hidden;width:15px;height:15px"><img src="https://lh3.googleusercontent.com/vLSCYEGN7oLtmhaSdBjzq9psnzQaLlY-fa7QgcbvWWjPvwrwCDr2qX7nHFglSmQHxwPPf0DmH17j6TgttmB54ke2L4x7BJYp3DiNISF5do3G2gsBdS1v9_KchSJAc-K_dh6FtGUxJtHmrOtZPDPZEnw" width="15" height="15" style="margin-left:0px;margin-top:0px" crossorigin="" class="CToWUd" data-bit="iit"></span></span><span style="font-size:8pt;font-family:Arial,sans-serif;color:rgb(0,0,0);vertical-align:baseline"> </span><a href="https://www.multisyscorp.com/" target="_blank" data-saferedirecturl="https://www.google.com/url?q=https://www.multisyscorp.com/&amp;source=gmail&amp;ust=1731657202461000&amp;usg=AOvVaw0G3eJRh1oAS1IGmlpIrvQt"><span style="font-size:10pt;font-family:Arial,sans-serif;color:rgb(17,85,204);vertical-align:baseline">www.multisyscorp.com</span></a></p></td></tr><tr style="height:30pt"><td style="vertical-align:top;background-color:rgb(249,249,249);overflow:hidden"><p dir="ltr" style="line-height:1.2;margin-top:0pt;margin-bottom:0pt"><span style="font-size:9pt;font-family:Arial,sans-serif;color:rgb(0,0,0);vertical-align:baseline"><span style="border:none;display:inline-block;overflow:hidden;width:121px;height:27px"><img src="https://lh4.googleusercontent.com/m4vMFXp2L8_GL568U8By8vFWFacVg-4s_RAjIFlQMJJvTQfgik58VwVIJY8g1zW8g9n-__YYXwDYO9kZzDidrTIWrRieExlSnwvtRuEqikp5XgZ3G9xAyIXd8eFN-k42XJRXhK0APe5u1FD8si56y44" width="121" height="27" style="margin-left:0px;margin-top:0px" crossorigin="" class="CToWUd" data-bit="iit"></span></span></p></td></tr></tbody></table></div><br><p dir="ltr" style="line-height:1.38;margin-top:8pt;margin-bottom:0pt"><span style="font-size:11pt;color:rgb(0,0,0);background-color:transparent;vertical-align:baseline;white-space:pre-wrap"><font face="times new roman, serif">DISCLAIMER:</font></span></p><p dir="ltr" style="line-height:1.38;margin-top:8pt;margin-bottom:0pt"><span style="font-size:11pt;color:rgb(0,0,0);background-color:transparent;vertical-align:baseline;white-space:pre-wrap"><font face="times new roman, serif">The content of this email is confidential and intended solely for the use of the individual or entity to whom it is addressed. If you received this email in error, please notify the sender or system manager. It is strictly forbidden to disclose, copy or distribute any part of this message.</font></span></p></span></div></div></div>`;

const storeSheetUrl = "https://docs.google.com/spreadsheets/d/1TJ10XqwS_cTQfkxKKWJaE5zhBdE2pDgIdZ_Zw9JQD_U";
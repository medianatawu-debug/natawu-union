/** NATAWU BACKEND - Google Apps Script
 * 1) Create a Google Sheet for NATAWU applications.
 * 2) Extensions -> Apps Script.
 * 3) Paste this file and set CONFIG.SPREADSHEET_ID.
 * 4) Deploy -> New deployment -> Web app -> Execute as Me -> Anyone.
 * 5) Put the Web App URL into apply-membership.html and js/main.js.
 *
 * Membership applications and website enquiries are written to two
 * separate, separately-formatted sheet tabs (see CONFIG below).
 * The Applications tab includes a live "Applications by Province"
 * dashboard (formulas, so it updates itself as rows are added) and a
 * NATAWU letterhead logo at the top of both tabs for printing.
 * Both notification emails are sent as branded HTML, not plain text.
 *
 * NOTE ON FILE SIZE: most of this file's length is the NATAWU logo,
 * stored below as a base64-encoded image (LOGO_BASE64_CHUNKS). That's
 * expected — don't remove or edit that block.
 */
const CONFIG={
  SPREADSHEET_ID:'PASTE_GOOGLE_SHEET_ID_HERE',
  APPLICATIONS_SHEET:'Applications',
  ENQUIRIES_SHEET:'Enquiries',
  NATAWU_EMAIL:'sgift8083@gmail.com',
  MEDIA_EMAIL:'media.natawu@gmail.com'
};

const BRAND={
  green:'#123B27',
  greenDark:'#081F14',
  gold:'#D6AD55',
  cream:'#F7F4ED',
  ink:'#17211B',
  inkSoft:'#647168',
  border:'#E1D6BE'
};

const PROVINCES=['Eastern Cape','Free State','Gauteng','KwaZulu-Natal','Limpopo','Mpumalanga','Northern Cape','North West','Western Cape'];

const MEMBERSHIP_FIELDS=['applicationNumber','submittedAt','province','localOffice','surname','names','fullName','placeOfWork','unionSector','idNumber','passportNumber','dateOfBirth','gender','disabled','contactNumber','employeeNumber','grossSalary','department','position','organiserAdmin','organiserAdminCode','organiserAdminContact','shopSteward','shopStewardContact','nextOfKinName','nextOfKinContact','companyName','companyRefNumber','companyPostalAddress','companyCode','companyContactNumber','companyFaxNumber','companyContactPerson','companyEmail','memberAgreement','witnessName','stopOrderDate','membershipType','currentUnion','currentUnionMembershipNumber','resignationDeclarationAccepted','resignationDate','termsAccepted','popiaConsent'];
const ENQUIRY_FIELDS=['reference','submittedAt','firstName','lastName','email','phone','branch','message'];

/* =====================================================================
   ENTRY POINT
   ===================================================================== */
function doPost(e){
  try{
    const data=JSON.parse(e.postData.contents||'{}');
    if(data.action==='contact') return json(handleContact(data));
    return json(handleMembership(data));
  }catch(err){return json({success:false,message:err.message});}
}
function json(obj){return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);}

function getSheet(name){
  const ss=SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  let sh=ss.getSheetByName(name);
  if(!sh) sh=ss.insertSheet(name);
  return sh;
}

/* =====================================================================
   SUBMISSION HANDLERS
   ===================================================================== */
function handleMembership(data){
  const sh=getSheet(CONFIG.APPLICATIONS_SHEET);
  buildApplicationsLayout(sh);
  const ref='NATAWU-'+Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyyMMdd-HHmmss');
  const now=new Date();
  data.applicationNumber=ref; data.submittedAt=now.toISOString();
  sh.appendRow(MEMBERSHIP_FIELDS.map(k=>data[k]??''));
  sendMembershipEmail(data,ref,now);
  return {success:true,applicationNumber:ref};
}

function handleContact(data){
  const sh=getSheet(CONFIG.ENQUIRIES_SHEET);
  buildEnquiriesLayout(sh);
  const now=new Date();
  const ref='WEB-'+Utilities.formatDate(now,Session.getScriptTimeZone(),'yyyyMMdd-HHmmss');
  sh.appendRow(ENQUIRY_FIELDS.map(k=>k==='reference'?ref:k==='submittedAt'?now.toISOString():(data[k]||'')));
  sendEnquiryEmail(data,ref,now);
  return {success:true,reference:ref};
}

/* Run this once manually from the Apps Script editor (function dropdown
   -> setup -> Run) to pre-create both formatted tabs. Safe to run more
   than once — it only builds a tab's layout if that tab is still empty. */
function setup(){
  buildApplicationsLayout(getSheet(CONFIG.APPLICATIONS_SHEET));
  buildEnquiriesLayout(getSheet(CONFIG.ENQUIRIES_SHEET));
}

/* =====================================================================
   SHEET LAYOUT (logo letterhead + live province dashboard + headers)
   ===================================================================== */
function columnToLetter(col){
  let letter='';
  while(col>0){
    const rem=(col-1)%26;
    letter=String.fromCharCode(65+rem)+letter;
    col=Math.floor((col-1)/26);
  }
  return letter;
}

function insertLogo(sh,row,col,size){
  const img=sh.insertImage(getLogoBlob(),col,row,4,4);
  img.setWidth(size).setHeight(size);
  return img;
}

function buildApplicationsLayout(sh){
  if(sh.getLastRow()>0) return; // already built — never rebuild over real data
  const n=MEMBERSHIP_FIELDS.length;
  sh.setColumnWidths(1,n,170);

  sh.getRange(1,3,1,12).merge().setValue('NATAWU — MEMBERSHIP APPLICATIONS')
    .setFontWeight('bold').setFontSize(17).setFontColor(BRAND.green).setVerticalAlignment('middle');
  sh.getRange(2,3,1,12).merge().setValue('Live record of all online membership applications')
    .setFontStyle('italic').setFontColor(BRAND.inkSoft).setFontSize(10).setVerticalAlignment('middle');
  sh.setRowHeight(1,44);
  sh.setRowHeight(2,22);
  insertLogo(sh,1,1,58);

  sh.getRange(4,1,1,n).merge().setValue('APPLICATIONS BY PROVINCE')
    .setFontWeight('bold').setFontColor('#FFFFFF').setBackground(BRAND.green).setFontSize(11)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.setRowHeight(4,26);

  const provinceHeaderRange=sh.getRange(5,1,1,PROVINCES.length);
  provinceHeaderRange.setValues([PROVINCES])
    .setFontWeight('bold').setBackground(BRAND.cream).setFontColor(BRAND.green)
    .setHorizontalAlignment('center').setFontSize(9)
    .setBorder(true,true,true,true,true,true,BRAND.border,SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(5,22);

  const provinceCol=MEMBERSHIP_FIELDS.indexOf('province')+1;
  const colLetter=columnToLetter(provinceCol);
  const formulas=PROVINCES.map(function(p){return '=COUNTIF('+colLetter+'9:'+colLetter+'5000,"'+p+'")';});
  const countRange=sh.getRange(6,1,1,PROVINCES.length);
  countRange.setFormulas([formulas])
    .setFontWeight('bold').setFontSize(15).setFontColor(BRAND.green).setHorizontalAlignment('center')
    .setBorder(true,true,true,true,true,true,BRAND.border,SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(6,32);

  sh.getRange(8,1,1,n).setValues([MEMBERSHIP_FIELDS])
    .setFontWeight('bold').setFontColor('#FFFFFF').setBackground(BRAND.green).setFontSize(11)
    .setVerticalAlignment('middle').setHorizontalAlignment('left');
  sh.setRowHeight(8,36);

  sh.setFrozenRows(8);

  try{
    sh.getRange(9,1,300,n).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY,false,false);
  }catch(err){ /* formatting is best-effort; never block a real submission */ }
}

function buildEnquiriesLayout(sh){
  if(sh.getLastRow()>0) return;
  const n=ENQUIRY_FIELDS.length;
  sh.setColumnWidths(1,n,170);

  sh.getRange(1,3,1,6).merge().setValue('NATAWU — WEBSITE ENQUIRIES')
    .setFontWeight('bold').setFontSize(17).setFontColor(BRAND.green).setVerticalAlignment('middle');
  sh.getRange(2,3,1,6).merge().setValue('Live record of all contact form submissions')
    .setFontStyle('italic').setFontColor(BRAND.inkSoft).setFontSize(10).setVerticalAlignment('middle');
  sh.setRowHeight(1,44);
  sh.setRowHeight(2,22);
  insertLogo(sh,1,1,58);

  sh.getRange(4,1,1,n).setValues([ENQUIRY_FIELDS])
    .setFontWeight('bold').setFontColor('#FFFFFF').setBackground(BRAND.green).setFontSize(11)
    .setVerticalAlignment('middle').setHorizontalAlignment('left');
  sh.setRowHeight(4,36);

  sh.setFrozenRows(4);

  try{
    sh.getRange(5,1,300,n).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY,false,false);
  }catch(err){}
}

/* =====================================================================
   BRANDED HTML EMAILS
   ===================================================================== */
function escapeHtml(str){
  return String(str==null?'':str)
    .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')
    .replace(/"/g,'&quot;').replace(/'/g,'&#39;');
}
function fieldLabel(key){
  const s=key.replace(/([A-Z])/g,' $1').trim();
  return s.charAt(0).toUpperCase()+s.slice(1);
}
function displayValue(key,data){
  const boolFields=['memberAgreement','resignationDeclarationAccepted','termsAccepted','popiaConsent'];
  let val=data[key];
  if(boolFields.indexOf(key)!==-1) return val?'Agreed':'Not agreed';
  if(val===undefined||val===null||val==='') return '—';
  return val;
}
function htmlSectionTable(title,data,keys){
  const rows=keys.map(function(k,i){
    const bg=i%2===0?'#FFFFFF':BRAND.cream;
    return '<tr>'+
      '<td style="padding:9px 14px;font-size:12.5px;color:'+BRAND.inkSoft+';font-weight:600;background:'+bg+';border-bottom:1px solid '+BRAND.border+';width:44%;">'+escapeHtml(fieldLabel(k))+'</td>'+
      '<td style="padding:9px 14px;font-size:13px;color:'+BRAND.ink+';font-weight:700;background:'+bg+';border-bottom:1px solid '+BRAND.border+';">'+escapeHtml(displayValue(k,data))+'</td>'+
      '</tr>';
  }).join('');
  return '<tr><td style="padding:24px 24px 8px;">'+
      '<div style="font-size:14px;font-weight:800;color:'+BRAND.green+';border-bottom:2px solid '+BRAND.gold+';padding-bottom:6px;">'+escapeHtml(title)+'</div>'+
    '</td></tr>'+
    '<tr><td style="padding:0 24px 4px;">'+
      '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;border:1px solid '+BRAND.border+';border-radius:8px;overflow:hidden;">'+rows+'</table>'+
    '</td></tr>';
}

function buildMembershipEmailHtml(data,ref,now){
  const dateStr=Utilities.formatDate(now,Session.getScriptTimeZone(),'d MMM yyyy, HH:mm');
  const sections=[
    ['New Member Details',['province','localOffice','surname','names','placeOfWork','unionSector','idNumber','passportNumber','dateOfBirth','gender','disabled','contactNumber','employeeNumber','grossSalary','department','position','organiserAdmin','organiserAdminCode','organiserAdminContact','shopSteward','shopStewardContact','nextOfKinName','nextOfKinContact']],
    ['Company Details',['companyName','companyRefNumber','companyPostalAddress','companyCode','companyContactNumber','companyFaxNumber','companyContactPerson','companyEmail']],
    ['Stop Order Authorisation',['memberAgreement','witnessName','stopOrderDate']],
    ['Resignation From Current Union',['membershipType','currentUnion','currentUnionMembershipNumber','resignationDeclarationAccepted','resignationDate']],
    ['Agreements & Consent',['termsAccepted','popiaConsent']]
  ];
  const sectionsHtml=sections.map(function(s){return htmlSectionTable(s[0],data,s[1]);}).join('');
  return '<div style="background:'+BRAND.cream+';padding:28px 12px;font-family:Arial,Helvetica,sans-serif;">'+
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="max-width:640px;margin:0 auto;background:#FFFFFF;border-radius:14px;overflow:hidden;box-shadow:0 4px 18px rgba(18,59,39,0.14);">'+
      '<tr><td style="background:'+BRAND.green+';padding:26px 28px;">'+
        '<table role="presentation" cellpadding="0" cellspacing="0"><tr>'+
          '<td style="padding-right:14px;vertical-align:top;"><img src="cid:natawuLogo" width="48" height="48" style="display:block;border-radius:8px;background:#fff;padding:3px;" alt="NATAWU"></td>'+
          '<td>'+
            '<div style="font-size:11px;font-weight:800;letter-spacing:.08em;color:'+BRAND.gold+';text-transform:uppercase;">NATAWU Membership Application</div>'+
            '<div style="font-size:23px;font-weight:800;color:#FFFFFF;letter-spacing:.01em;margin-top:3px;">'+escapeHtml(String(data.fullName||'').toUpperCase())+'</div>'+
            '<div style="font-size:12.5px;color:#EDE6D6;margin-top:9px;">Application Number: <strong style="color:#FFFFFF;">'+escapeHtml(ref)+'</strong></div>'+
            '<div style="font-size:11.5px;color:#B9C9BE;margin-top:2px;">Submitted '+dateStr+'</div>'+
          '</td>'+
        '</tr></table>'+
      '</td></tr>'+
      sectionsHtml+
      '<tr><td style="padding:18px 24px 28px;">'+
        '<p style="font-size:11.5px;color:'+BRAND.inkSoft+';margin:0;">This application was submitted via the NATAWU website and recorded automatically in the Applications sheet. No action is needed to acknowledge this email.</p>'+
      '</td></tr>'+
    '</table>'+
  '</div>';
}

function buildMembershipEmailPlainText(data,ref,now){
  let body='NATAWU MEMBERSHIP APPLICATION\n\nApplication Number: '+ref+'\nSubmitted: '+now+'\n\n';
  MEMBERSHIP_FIELDS.slice(2).forEach(function(k){body+=fieldLabel(k)+': '+(data[k]??'')+'\n';});
  return body;
}

function sendMembershipEmail(data,ref,now){
  const subject='Application - '+(data.fullName||'NATAWU Member');
  MailApp.sendEmail({
    to:CONFIG.NATAWU_EMAIL,
    subject:subject,
    body:buildMembershipEmailPlainText(data,ref,now),
    htmlBody:buildMembershipEmailHtml(data,ref,now),
    inlineImages:{natawuLogo:getLogoBlob()}
  });
}

function buildEnquiryEmailHtml(data,ref,now){
  const dateStr=Utilities.formatDate(now,Session.getScriptTimeZone(),'d MMM yyyy, HH:mm');
  const rows=[['Email',data.email],['Phone',data.phone||'—'],['Branch',data.branch||'—']]
    .map(function(pair,i){
      const bg=i%2===0?'#FFFFFF':BRAND.cream;
      return '<tr>'+
        '<td style="padding:9px 14px;font-size:12.5px;color:'+BRAND.inkSoft+';font-weight:600;background:'+bg+';border-bottom:1px solid '+BRAND.border+';width:34%;">'+escapeHtml(pair[0])+'</td>'+
        '<td style="padding:9px 14px;font-size:13px;color:'+BRAND.ink+';font-weight:700;background:'+bg+';border-bottom:1px solid '+BRAND.border+';">'+escapeHtml(pair[1])+'</td>'+
        '</tr>';
    }).join('');
  return '<div style="background:'+BRAND.cream+';padding:28px 12px;font-family:Arial,Helvetica,sans-serif;">'+
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="max-width:600px;margin:0 auto;background:#FFFFFF;border-radius:14px;overflow:hidden;box-shadow:0 4px 18px rgba(18,59,39,0.14);">'+
      '<tr><td style="background:'+BRAND.green+';padding:24px 26px;">'+
        '<table role="presentation" cellpadding="0" cellspacing="0"><tr>'+
          '<td style="padding-right:14px;vertical-align:top;"><img src="cid:natawuLogo" width="42" height="42" style="display:block;border-radius:8px;background:#fff;padding:3px;" alt="NATAWU"></td>'+
          '<td>'+
            '<div style="font-size:11px;font-weight:800;letter-spacing:.08em;color:'+BRAND.gold+';text-transform:uppercase;">NATAWU Website Enquiry</div>'+
            '<div style="font-size:20px;font-weight:800;color:#FFFFFF;margin-top:3px;">'+escapeHtml(((data.firstName||'')+' '+(data.lastName||'')).trim().toUpperCase())+'</div>'+
            '<div style="font-size:11.5px;color:#B9C9BE;margin-top:9px;">Reference '+escapeHtml(ref)+' &middot; '+dateStr+'</div>'+
          '</td>'+
        '</tr></table>'+
      '</td></tr>'+
      '<tr><td style="padding:22px 24px 4px;">'+
        '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;border:1px solid '+BRAND.border+';border-radius:8px;overflow:hidden;">'+rows+'</table>'+
      '</td></tr>'+
      '<tr><td style="padding:18px 24px 26px;">'+
        '<div style="font-size:12px;font-weight:800;color:'+BRAND.green+';text-transform:uppercase;letter-spacing:.04em;margin-bottom:6px;">Message</div>'+
        '<div style="font-size:13.5px;color:'+BRAND.ink+';line-height:1.6;background:'+BRAND.cream+';border:1px solid '+BRAND.border+';border-radius:8px;padding:14px 16px;white-space:pre-wrap;">'+escapeHtml(data.message||'(no message provided)')+'</div>'+
      '</td></tr>'+
    '</table>'+
  '</div>';
}

function sendEnquiryEmail(data,ref,now){
  const subject='NATAWU Website Enquiry - '+(data.firstName||'')+' '+(data.lastName||'');
  const plainBody='New website enquiry\n\nReference: '+ref+'\nName: '+data.firstName+' '+data.lastName+'\nEmail: '+data.email+'\nPhone: '+(data.phone||'')+'\nBranch: '+(data.branch||'')+'\n\nMessage:\n'+(data.message||'');
  MailApp.sendEmail({
    to:CONFIG.NATAWU_EMAIL,
    cc:CONFIG.MEDIA_EMAIL,
    subject:subject,
    body:plainBody,
    htmlBody:buildEnquiryEmailHtml(data,ref,now),
    inlineImages:{natawuLogo:getLogoBlob()}
  });
}

/* =====================================================================
   NATAWU LOGO (base64-encoded PNG, used for the sheet letterhead and
   the inline email logo). Do not edit this block by hand.
   ===================================================================== */
const LOGO_BASE64_CHUNKS=[
  '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'
];
function getLogoBlob(){
  return Utilities.newBlob(Utilities.base64Decode(LOGO_BASE64_CHUNKS.join('')),'image/png','natawu-logo.png');
}

/* SPDX-License-Identifier: AGPL-3.0-only */
(function(root){
'use strict';
function passwordError(){var e=new Error('Invalid document password options');e.code='password-options';return e;}
 function passwordOption(value){
  if(value===undefined)return;
  if(typeof value!=='string'||!value.length||value.length>255||/[\x00-\x08\x0b\x0c\x0e-\x1f\ufffe\uffff]/.test(value))throw passwordError();
  for(var i=0;i<value.length;i++){
   var code=value.charCodeAt(i);
   if(code>=0xd800&&code<=0xdbff){var next=value.charCodeAt(++i);if(!(next>=0xdc00&&next<=0xdfff))throw passwordError();}
   else if(code>=0xdc00&&code<=0xdfff)throw passwordError();
  }
 }

function xml(value){return value.replace(/[&<>"'\r]/g,function(c){return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&apos;','\r':'&#13;'}[c];});}
function check(job){passwordOption(job.openPassword);passwordOption(job.savePassword);if(job.savePassword!==undefined&&['docx','xlsx','pptx','odt','ods','odp'].indexOf(job.to)<0)throw passwordError();}
function tags(job){return(job.openPassword===undefined?'':'<m_sPassword>'+xml(job.openPassword)+'</m_sPassword>')+(job.savePassword===undefined?'':'<m_sSavePassword>'+xml(job.savePassword)+'</m_sSavePassword>');}
function nativeError(status){var e=new Error(status===90?'A password is required to open this document':'The document password is incorrect');e.code=status===90?'password-required':'password-incorrect';return e;}
function failure(e,job,status){var typed=e.code&&/^password-/.test(e.code);return {error:typed?e.message:job.openPassword!==undefined||job.savePassword!==undefined?'Local document conversion failed':String(e.message||e),code:typed?e.code:typeof status==='number'?'conversion':'runtime',nativeStatus:status};}
root.EOWasmPasswordOptions={check:check,tags:tags,nativeError:nativeError,failure:failure};
})(self);

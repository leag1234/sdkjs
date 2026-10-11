/*
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */
(function (window) {
 'use strict';
 var script = document.currentScript.src, base = new URL('.', script).href;
 var formats = {docx:65,odt:67,pptx:129,odp:131,xlsx:257,ods:259,doct:8193,xlst:8194,pptt:8195};
 function error(code,message) { var e=new Error(message); e.code=code; return e; }
 async function convert(input,from,to,options) {
  options=options||{};
  if (!formats[from] || !formats[to]) throw error('format','Unsupported conversion format');
  if (!(input instanceof Uint8Array) && !(input instanceof ArrayBuffer)) throw error('input','Expected document bytes');
  if (options.signal && options.signal.aborted) throw error('cancelled','Conversion cancelled');
  var runtime=new URL(options.runtime || 'runtime/528e920e4f34-98317b47cc65-st/',base);
  if (!runtime.pathname.endsWith('/')) runtime.pathname += '/';
  if (runtime.origin !== window.location.origin) throw error('runtime','Conversion runtime must be same-origin');
  var bytes=new Uint8Array(input instanceof ArrayBuffer?input.slice(0):input.slice()),files=[];
  if (!bytes.length) throw error('input','Document is empty');
  var fonts=window.EOWasmFonts;
  if (!fonts) throw error('font','Shared font service is unavailable');
  var preparation=fonts.prepareFS({FS:{mkdirTree:function(){},writeFile:function(path,data){files.push({path:path,bytes:data.slice()});}}},options.fontNames||[]);
  var prepared=await new Promise(function(resolve,reject){
   function cancel(){cleanup();reject(error('cancelled','Conversion cancelled'));}
   function cleanup(){if(options.signal)options.signal.removeEventListener('abort',cancel);}
   if(options.signal)options.signal.addEventListener('abort',cancel,{once:true});
   preparation.then(function(result){cleanup();resolve(result);},function(e){cleanup();reject(e);});
  });
  if (options.signal && options.signal.aborted) throw error('cancelled','Conversion cancelled');
  return new Promise(function(resolve,reject){
   var worker=new Worker(new URL('worker.js',base)),done=false;
   function finish(err,result) {
    if(done)return;done=true;worker.terminate();
    if(options.signal)options.signal.removeEventListener('abort',cancel);
    err?reject(err):resolve(result);
   }
   function cancel(){finish(error('cancelled','Conversion cancelled'));}
   if(options.signal)options.signal.addEventListener('abort',cancel,{once:true});
   worker.onerror=function(e){finish(error('worker',e.message||'Conversion worker failed'));};
   worker.onmessage=function(e){
    var data=e.data;
    if(data.progress){try{if(options.onProgress)options.onProgress(data.progress);}catch(e){finish(error('progress',e.message));}return;}
    if(data.error){var failure=error(data.code||'conversion',data.error);failure.nativeStatus=data.nativeStatus;finish(failure);return;}
    var output=new Uint8Array(data.bytes);
    output.resources=data.resources;output.stats=Object.assign({},data.stats,{fontBytes:prepared.bytes,fontFiles:files.length-1});
    finish(null,output);
   };
   try{worker.postMessage({bytes:bytes,from:from,to:to,runtime:runtime.href,files:files,resources:options.resources||[]},[bytes.buffer].concat(files.map(function(f){return f.bytes.buffer;})));}catch(e){finish(error('input','Conversion input cannot be sent to a worker'));}
  });
 }
 window.EOWasmConvert={convert:convert,formats:formats};
})(window);

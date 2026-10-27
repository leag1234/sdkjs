/*
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */
(function(window){
 'use strict';
 var base=new URL('../clientpdf/',document.currentScript.src).href,loading;
 function script(name){return new Promise(function(resolve,reject){
  var element=document.createElement('script');element.src=base+name;
  element.onload=resolve;element.onerror=function(){reject(new Error('Cannot load the bundled PDF exporter'));};
  document.head.appendChild(element);
 });}
 function cancelled(signal){if(signal&&signal.aborted)throw new Error('Conversion cancelled');}
 async function libraries(){
  if(!loading)loading=(async function(){
   if(!AscCommon.ClientPdf)await script('clientpdf.js');
   if(!window.EOWasmPdf)await script('wasm-pdf.js');
  })().catch(function(e){loading=null;throw e;});
  return loading;
 }
 window.EOLocalPdf=async function(api,format,options,signal){
  cancelled(signal);await libraries();cancelled(signal);
  if(format==='docx'||format==='odt'){
   var result=await AscCommon.ClientPdf.build(api,options);
   cancelled(signal);return result;
  }
  var response=await fetch(base+'runtime/current.json');
  if(!response.ok)throw new Error('Bundled PDF runtime missing');
  var runtime=await response.json();
  if(!/^[a-f0-9-]+st$/.test(runtime.artifact_version))throw new Error('Invalid bundled PDF runtime');
  var writer=new EOWasmPdf.Writer({getRawImage:AscCommon.ClientPdf.getRawImage,
   wasmRuntime:base+'runtime/'+runtime.artifact_version+'/'});
  var memory;
  if(format==='xlsx'||format==='ods'){
   var adjust=options.advancedOptions;
   if(!adjust){adjust=new Asc.asc_CAdjustPrint();adjust.asc_setPrintType(Asc.c_oAscPrintType.EntireWorkbook);}
   api.wbModel.dependencyFormulas.calcTree();
   api.wb._executeWithoutZoom(function(){
    var previous=api.wb.printPreviewState.advancedOptions;
    api.wb.printPreviewState.advancedOptions=adjust;
    try{memory=api.wb.printSheets(api.wb.calcPagesPrint(adjust),null,adjust).DocumentRenderer.Memory;}
    finally{api.wb.printPreviewState.advancedOptions=previous;}
   });
  }else{
   var renderer=new AscCommon.CDocumentRenderer();
   renderer.InitPicker(AscCommon.g_oTextMeasurer.m_oManager);
   renderer.VectorMemoryForPrint=new AscCommon.CMemory();renderer.IsNoDrawingEmptyPlaceholder=true;
   renderer.DocInfo(api.asc_getCoreProps());
   var marks=api.ShowParaMarks,count=api.asc_nativePrintPagesCount(),native=options.nativeOptions||{};
   var ranges=native.pages!==undefined?AscCommon.getNativePrintRanges(native.pages,native.currentPage,count):undefined;
   api.ShowParaMarks=false;
   try{for(var i=0;i<count;i++){cancelled(signal);if(!ranges||ranges[i])api.asc_nativePrint(renderer,i,{});}}
   finally{api.ShowParaMarks=marks;}
   memory=renderer.Memory;
  }
  var data=memory.GetDataUint8(0,memory.GetCurPosition()),bytes,backend='wasm',fallbacks;
  cancelled(signal);
  try{await writer.feed(data);bytes=await writer.finish();}
  catch(error){
   if(!error||error.code!=='cjk')throw error;
   await AscCommon.ClientPdf.loadLibs();cancelled(signal);
   writer=new EOPdf.Writer({fontkit:window.fontkit,getFont:AscCommon.ClientPdf.fonts.get,
    getImage:AscCommon.ClientPdf.getImage,signal:signal,producer:'ElixirOffice',creator:'ElixirOffice'});
   await writer.feed(data);bytes=await writer.finish();backend='js';
   fallbacks=[{from:'wasm',to:'js',reason:'cjk',detail:error.detail||error.message}];
  }
  cancelled(signal);
  return {bytes:bytes,stats:Object.assign({},writer.stats,{backend:backend},fallbacks?{fallbacks:fallbacks}:{})};
 };
})(window);

// SPDX-License-Identifier: AGPL-3.0-only
"use strict";
// Table cell borders in the pptx binary writer: the "No Border" preset of the table settings sends
// Value 0 (AscWord.BorderType.nil). That border is not drawn, so it must be written as a:noFill and
// never as a solid line (it used to be saved as a black 0.5 pt line that reappeared on reopen).
const {test}=require("node:test"),assert=require("node:assert/strict"),fs=require("node:fs"),path=require("node:path"),vm=require("node:vm");
const src=fs.readFileSync(path.resolve(__dirname,"../../common/Shapes/SerializeWriter.js"),"utf8");
function method(marker){
 const a=src.indexOf(marker);assert(a>=0,marker);
 const b=src.indexOf("\n    this.",a+marker.length);assert(b>a);
 return src.slice(a+marker.length,b).trim().replace(/;$/,"");
}
function run(marker,border){
 const log=[];
 const oThis={StartRecord:t=>log.push("start"+t),EndRecord:()=>log.push("end"),WriteUChar:()=>{},
  _WriteInt2:()=>log.push("size"),WriteRecord2:(t,fill)=>log.push(fill&&fill.fill?fill.fill.kind:"nofill?"),
  WriteRecord3:()=>{},_WriteBool1:()=>{},_WriteInt1:()=>{}};
 const sandbox={oThis,log,border_None:-1,AscWord:{BorderType:{none:-1,nil:0,single:1}},
  g_nodeAttributeStart:0,g_nodeAttributeEnd:1,
  AscFormat:{CUniFill:function(){},CNoFill:function(){this.kind="NoFill"}}};
 const fn=vm.runInNewContext("("+method(marker)+")",sandbox);
 fn(border);return log;
}
const M="this.WriteTableCellBorder = ";
test("Value 0 and Value -1 are written as a:noFill without a width",()=>{
 for(const v of [0,-1]){
  const log=run(M,{Value:v,Size:0.176,Unifill:{fill:{}}});
  assert(log.includes("NoFill"),"Value "+v);
  assert(!log.includes("size"),"no width for Value "+v);
 }
});
test("a single border keeps its width",()=>{
 const log=run(M,{Value:1,Size:0.176,Unifill:null});
 assert(log.includes("size"));assert(!log.includes("NoFill"));
});
test("the table style border writer treats Value 0 as none as well",()=>{
 assert(/WriteTableCellBorderLineStyle = function[\s\S]{0,200}BorderType\.nil/.test(src));
});

// Reader side: a:lnX with a:noFill must come back as "no border", not as a "single" border with a no-fill colour.
const fmt=fs.readFileSync(path.resolve(__dirname,"../../common/Drawings/Format/Format.js"),"utf8");
function fillDocumentBorder(line){
 const a=fmt.indexOf("CLn.prototype.fillDocumentBorder = function(oBorder) {");assert(a>=0);
 const b=fmt.indexOf("CLn.prototype.fromDocumentBorder",a);assert(b>a);
 const text=fmt.slice(a,b).trim().replace(/^CLn\.prototype\.fillDocumentBorder = /,"").replace(/;$/,"");
 const sandbox={c_oAscFill:{FILL_TYPE_NOFILL:0,FILL_TYPE_SOLID:1},AscWord:{BorderType:{none:-1,nil:0,single:1}},AscCommonWord:{border_Single:1}};
 const border={};vm.runInNewContext("("+text+")",sandbox).call(line,border);return border;
}
test("a no-fill line reads as border none and keeps its width",()=>{
 const b=fillDocumentBorder({Fill:{fill:{type:0}},w:12700});
 assert.equal(b.Value,-1);assert.equal(b.Size,12700/36000);
});
test("a solid or default line still reads as a single border",()=>{
 assert.equal(fillDocumentBorder({Fill:{fill:{type:1}},w:12700}).Value,1);
 assert.equal(fillDocumentBorder({Fill:null,w:null}).Value,1);
});
test("the reader gives border none for a:noFill in cells and in table styles (both go through fillDocumentBorder)",()=>{
 const ser=fs.readFileSync(path.resolve(__dirname,"../../common/Shapes/Serialize.js"),"utf8");
 assert(/ReadTableBorderLn = function\(\)[\s\S]{0,200}fillDocumentBorder\(border\)/.test(ser));
 assert(/ReadTableBorderLineStyle = function[\s\S]{0,400}fillDocumentBorder\(_border\)/.test(ser));
});

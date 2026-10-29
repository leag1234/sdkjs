// SPDX-License-Identifier: AGPL-3.0-only
'use strict';
const {test}=require('node:test'),assert=require('node:assert/strict'),fs=require('node:fs'),path=require('node:path'),vm=require('node:vm');
const root=path.resolve(__dirname,'../..');
const serial=fs.readFileSync(path.join(root,'word/Editor/Serialize2.js'),'utf8');
function method(source,marker,next){
 const start=source.indexOf(marker);assert(start>=0,marker);
 const end=source.indexOf(next,start+marker.length);assert(end>start,next);
 return source.slice(start+marker.length,end).trim().replace(/;$/,'');
}
const types={CompatSetting:1,Flags1:2,Flags2:3};
const sandbox={c_oSerCompat:types,c_oSerConstants:{ReadOk:0,ReadUnknown:1}};
const read=vm.runInNewContext('('+method(serial,'this.ReadCompat = ','this.ReadCompatSetting = ')+')',sandbox);
const write=vm.runInNewContext('('+method(serial,'this.WriteCompat = ','this.WriteCompatSetting = ')+')',sandbox);
function readFlags(flags){
 const settings={};read.call({stream:{GetULong:()=>flags}},types.Flags2,4,settings);return settings;
}
function writeFlags(settings){
 let type;const values={};
 write.call({Document:{GetCompatibilityMode:()=>15,GetDocumentSettings:()=>settings},saveParams:{isCompatible:true},
  WriteCompatSetting(){},bs:{WriteItem(t,fn){type=t;fn()}},memory:{WriteULong(value){values[type]=value>>>0}}});
 return values[types.Flags2];
}
test('explicit floating no-break bit survives x2t-compatible settings import and save',()=>{
 const flags=(1<<9)|(1<<17)|(1<<27),settings=readFlags(flags);
 assert.equal(settings.DoNotBreakWrappedTables,true);
 assert.equal(settings.UseFELayout,true);assert.equal(settings.SplitPageBreakAndParaMark,true);
 assert.equal(writeFlags(settings),flags);
});
test('false or absent no-break does not set the serialized bit or consume neighboring flags',()=>{
 for(const flags of [0,1<<8,1<<10]){
  const settings=readFlags(flags);
  assert.equal(settings.DoNotBreakWrappedTables,false);
  assert.equal(writeFlags(settings)&(1<<9),0);
 }
});
const docSource=fs.readFileSync(path.join(root,'word/Editor/Document.js'),'utf8');
const flowSource=docSource.slice(docSource.indexOf('CDocument.prototype.private_RecalculateFlowTable '),docSource.indexOf('CDocument.prototype.private_RecalculateFlowParagraph '));
const bits={next:1,page:2,cur:4,prev:8,last:16};
function fixture(noBreak,columns=1,first=true){
 const resets=[],saved={};let resetCount=0,can=first;
 const context={CDocument:function(){},CFlowTable:function(){},recalcresult_NextElement:bits.next,recalcresult_NextPage:bits.page,recalcresult_CurPage:bits.cur,recalcresult_PrevPage:bits.prev,recalcresultflags_LastFromNewPage:bits.last};
 vm.runInNewContext(flowSource,context);
 const element={PageNum:0,Set_DocumentIndex(){},Reset(...args){resets.push(args)},Recalculate_Page:()=>bits.page,IsContentOnFirstPage:()=>true};
 const recalc={Can_RecalcObject:()=>can,Check_FlowObject:o=>o===element,Set_FlowObject(_e,_p,_r,_c,info){Object.assign(saved,info)},Is_PageBreakBefore:()=>false,Reset(){resetCount++},FlowObjectPage:0,AdditionalInfo:{X:10,Y:20,XLimit:200,YLimit:180}};
 const document={RecalcInfo:recalc,Settings:noBreak===undefined?{}:{DoNotBreakWrappedTables:noBreak},GetDocumentSettings(){return this.Settings},
  Pages:[{Sections:[{IsCalculatingSectionBottomLine:()=>false}],AddFlowTable(){}}],DrawingObjects:{addFloatTable(){}}};
 const info={Element:element,X:10,Y:20,XLimit:200,YLimit:180,PageIndex:0,ColumnIndex:0,Index:1,StartIndex:0,ColumnsCount:columns,ResetStartElement:true,SectPr:{},SectionIndex:0,SectionIndexAbs:0,RecalcResult:0};
 const run=()=>context.CDocument.prototype.private_RecalculateFlowTable.call(document,info);
 return{run,info,resets,saved,setSecond(){can=false},resetCount:()=>resetCount};
}
for(const value of [undefined,false]){
 test('ordinary single-column float retains page limits and continuation: '+value,()=>{
  const f=fixture(value);f.run();assert.equal(f.resets[0][3],180);assert.equal(f.saved.YLimit,180);
  f.setSecond();f.run();assert.equal(f.info.RecalcResult,bits.page);assert.equal(f.resetCount(),0);
 });
}
test('explicit no-break uses the existing unbounded measure/reposition path only for this policy',()=>{
 const f=fixture(true);f.run();assert.equal(f.resets[0][3],10000);
 f.setSecond();f.run();assert.equal(f.info.RecalcResult,bits.next);assert.equal(f.resetCount(),1);
});
test('existing multicolumn whole-table path remains unchanged without the new flag',()=>{
 const f=fixture(false,2);f.run();assert.equal(f.resets[0][3],10000);
 f.setSecond();f.run();assert.equal(f.info.RecalcResult,bits.next);
});

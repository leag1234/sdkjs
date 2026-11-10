// SPDX-License-Identifier: AGPL-3.0-only
'use strict';
// Character borders (w:bdr, GitHub #66): TextPr.Border binary round trip, copy/merge/equality,
// the sdkjs <-> core type number and the merging of adjacent bordered elements.
const {test}=require('node:test'),assert=require('node:assert/strict'),fs=require('node:fs'),path=require('node:path'),vm=require('node:vm');
const root=path.resolve(__dirname,'../..');
const read=f=>fs.readFileSync(path.join(root,f),'utf8');
const styles=read('word/Editor/Styles.js'),serial=read('word/Editor/Serialize2.js'),para=read('word/Editor/Paragraph.js');
function method(source,marker,next){
 const start=source.indexOf(marker);assert(start>=0,marker);
 const end=source.indexOf(next,start+marker.length);assert(end>start,next);
 return source.slice(start+marker.length,end).trim().replace(/;$/,'');
}
class Border{
 constructor(v=0,size=0.5,space=0){this.Value=v;this.Size=size;this.Space=space;this.Color={r:0,g:0,b:0,Auto:true};this.Unifill=undefined}
 Copy(){const b=new Border(this.Value,this.Size,this.Space);b.Color={...this.Color};return b}
 IsNone(){return this.Value===0}
 IsEqual(o){return !!o&&this.Value===o.Value&&this.Size===o.Size&&this.Space===o.Space&&this.Color.Auto===o.Color.Auto}
 Is_Equal(o){return this.IsEqual(o)}
 Write_ToBinary(w){w.WriteByte(this.Value);w.WriteByte(this.Size*100);w.WriteByte(this.Space*100)}
 Read_FromBinary(r){this.Value=r.GetByte();this.Size=r.GetByte()/100;this.Space=r.GetByte()/100}
}
Border.FromObject=o=>Object.assign(new Border(),o);
const sandbox={CDocumentBorder:Border,IsEqualStyleObjects:(a,b)=>a===undefined&&b===undefined?true:(a===undefined||b===undefined?false:a.Is_Equal(b))};
const fn=(m,n)=>vm.runInNewContext('('+method(styles,m,n)+')',sandbox);
const write=fn('CTextPr.prototype.Write_ToBinary = ','CTextPr.prototype.Read_FromBinary = ');
const readBin=fn('CTextPr.prototype.Read_FromBinary = ','CTextPr.prototype.Check_NeedRecalc = ');
function stream(){const a=[];let p=0;return{a,
 GetCurPosition:()=>p,Skip:()=>{p++},Seek:n=>{p=n},WriteBool:v=>a[p++]=v,WriteByte:v=>a[p++]=v,WriteLong:v=>a[p++]=v,
 GetLong:()=>a[p++],GetByte:()=>a[p++],GetBool:()=>a[p++],setPos:n=>{p=n}}}
test('sdkjs rPr border type equals core c_oSerProp_rPrType::Border and is read and written',()=>{
 const m=/Border:\s*57/.exec(serial);assert(m,'Border: 57 in c_oSerProp_rPrType');
 assert(serial.includes('this.memory.WriteByte(c_oSerProp_rPrType.Border)'));
 assert(serial.includes('case c_oSerProp_rPrType.Border:'));
});
test('TextPr.Border survives Write_ToBinary/Read_FromBinary, also next to Ligatures',()=>{
 for(const lig of [undefined,3]){
  const tp={Ligatures:lig,Border:new Border(1,0.5,1)},s=stream();
  write.call(tp,s);s.setPos(0);
  const out={};readBin.call(out,s);
  assert.equal(out.Border.Value,1);assert.equal(out.Border.Size,0.5);assert.equal(out.Border.Space,1);
  assert.equal(out.Ligatures,lig);
 }
 const tp={Ligatures:2},s=stream();write.call(tp,s);s.setPos(0);const out={};readBin.call(out,s);
 assert.equal(out.Ligatures,2);assert.equal(out.Border,undefined);
 const none={},s2=stream();write.call(none,s2);assert.equal(s2.a[0]&(1<<31),0);
});
test('binary form without a border is byte-identical to the old format (old files read the same)',()=>{
 for(const lig of [0,1,3,15]){
  const s=stream();write.call({Ligatures:lig},s);
  assert.deepEqual(s.a,[1<<31|0,lig].map((v,i)=>i?v:s.a[0]));assert.equal(s.a[0]>>>31,1);assert.equal(s.a.length,2);
  s.setPos(0);const out={};readBin.call(out,s);assert.equal(out.Ligatures,lig);assert.equal(out.Border,undefined);
 }
 // a stream written by the old code: flags with bit 31 only, then the Ligatures byte
 const old=stream();old.a.push(1<<31,2);old.setPos(0);const out={};readBin.call(out,old);
 assert.equal(out.Ligatures,2);assert.equal(out.Border,undefined);
 // no Ligatures and no Border: no trailing byte at all
 const none=stream();write.call({},none);assert.equal(none.a.length,1);
});
test('TextPr.Border takes part in equality and diff',()=>{
 const a={Border:new Border(1)},b={Border:new Border(1)};
 assert.equal(sandbox.IsEqualStyleObjects(a.Border,b.Border),true);
 assert.equal(sandbox.IsEqualStyleObjects(a.Border,undefined),false);
 assert.equal(sandbox.IsEqualStyleObjects(a.Border,new Border(1,1)),false);
 assert(/IsEqualStyleObjects\(this\.Border, TextPr\.Border\)/.test(styles));
 assert(/oResultTextPr\.Border = this\.Border\.Copy\(\)/.test(styles));
});
test('Copy, Merge and Apply carry the border (style merge path)',()=>{
 for(const m of ['CTextPr.prototype.Copy = ','CTextPr.prototype.Merge = ','CTextPr.prototype.Apply = ','CTextPr.prototype.Set_FromObject = '])
  assert(method(styles,m,'\nCTextPr.prototype.').includes('Border'),m);
});
// merging of adjacent bordered elements (Paragraph.js CParaDrawingRangeLines)
const ctx={Math};
vm.runInNewContext(para.slice(para.indexOf('function CParaDrawingRangeLinesElement'),para.indexOf('function CParaDrawingRangeHorizontalLines'))
 +';this.L=CParaDrawingRangeLines;',ctx);
function stretch(borders){ // [[x0,x1,border,left,right]]
 const l=new ctx.L();
 for(const [x0,x1,b,L,R] of borders)l.Add(0,10,x0,x1,0.2,0,0,0,{Border:b,Left:L,Right:R});
 const out=[];let e;while((e=l.Get_Next()))out.push(e);return out.reverse();
}
test('adjacent elements with equal borders form one box; unequal or interrupted ones do not',()=>{
 const b1=new Border(1),b1b=new Border(1),b2=new Border(1,1);
 let r=stretch([[0,5,b1,true,false],[5,9,b1b,false,true]]);
 assert.equal(r.length,1);assert.equal(r[0].x0,0);assert.equal(r[0].x1,9);
 assert.equal(r[0].Additional.Left,true);assert.equal(r[0].Additional.Right,true);
 r=stretch([[0,5,b1,true,true],[5,9,b2,true,true]]);assert.equal(r.length,2);
 r=stretch([[0,5,b1,true,false],[5,9,b1,false,true],[12,15,b1,true,true]]);
 assert.equal(r.length,2);assert.equal(r[0].x1,9);assert.equal(r[1].x0,12);
});

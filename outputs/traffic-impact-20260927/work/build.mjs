import fs from 'node:fs/promises';
import {Workbook, SpreadsheetFile} from '@oai/artifact-tool';
const out = new URL('../', import.meta.url).pathname.replace(/^\/([A-Za-z]:)/,'$1');
const wb=Workbook.create();
const a=wb.worksheets.add('Traffic analysis'), l=wb.worksheets.add('SW added lane'), s=wb.worksheets.add('Inputs');
const navy='#243E5A', pale='#EAF0F6';
function val(sh,cell,x){sh.getRange(cell).values=[[x]];}
function fx(sh,cell,x){sh.getRange(cell).formulas=[[x]];}
function row(sh,n,arr){sh.getRangeByIndexes(n-1,0,1,arr.length).values=[arr];}
function head(sh,n,arr){row(sh,n,arr);let r=sh.getRangeByIndexes(n-1,0,1,arr.length);r.format.fill=navy;r.format.font={bold:true,color:'#FFFFFF'};r.format.wrapText=true;r.format.horizontalAlignment='center';r.format.rowHeight=43;}
function title(sh,cell,text){val(sh,cell,text);sh.getRange(cell).format.font={size:15,bold:true,color:'#000000'};}
function section(sh,n,text){val(sh,'A'+n,text);sh.getRange('A'+n).format.font={bold:true,size:11};sh.getRange('A'+n).format.rowHeight=28;}
function band(sh,start,end,cols){for(let i=start;i<=end;i++)if((i-start)%2===0)sh.getRangeByIndexes(i-1,0,1,cols).format.fill=pale;}
for(const [sh,range] of [[a,'A1:L46'],[l,'A1:J29'],[s,'A1:J42']]){
 sh.showGridLines=false;sh.getRange(range).format.font={name:'Arial',size:10,color:'#202B36'};
 sh.getRange(range).format.rowHeight=24;sh.getRange(range).format.columnWidth=14;
 sh.getRange(range).format.verticalAlignment='center';sh.getRange(range).setNumberFormat('#,##0.00');
}
a.tabColor=navy;l.tabColor=navy;s.tabColor='#7F9BB7';
title(s,'A2','Assignment inputs');
val(s,'A3','Source: Assignment 3 Project evaluation part I.docx, scenario, Questions 2 and 7–10, and Appendices A–C.');
section(s,5,'Appendix A per lane');head(s,6,['Demand (veh/h)','Flow (veh/h)','Speed (mph)','LOS']);
const speeds=[60,59,58,57,56,55,54,53,52,51,50,48,45,42,39,35,32,29,26,23,20];
const lookup=Array.from({length:21},(_,i)=>[1000+i*100,i<=10?1000+i*100:2000-(i-10)*100,speeds[i],i<5?'C':i<9?'D':i<11?'E':'F']);
s.getRange('A7:D27').values=lookup;s.getRange('A7:C27').setNumberFormat('#,##0');band(s,7,27,4);
val(s,'F5','Scenario and valuation assumptions');s.getRange('F5').format.font.bold=true;
s.getRange('F6:H6').values=[['Input','Value','Unit']];s.getRange('F6:H6').format.fill=navy;s.getRange('F6:H6').format.font.color='#FFFFFF';
const assumptions=[['Office area',800000,'sq ft'],['Trip rate area unit',1000,'sq ft'],['AM trip rate',1.45,'trips / 1,000 sq ft'],['AM inbound share',.87,'share'],['PM trip rate',1.5,'trips / 1,000 sq ft'],['PM outbound share',.85,'share'],['SW length',8,'miles'],['Added lanes',1,'peak direction'],['AM peak hours',2,'hours / day'],['PM peak hours',2,'hours / day'],['Value of time',30,'USD / vehicle-hour'],['Travel days',350,'days / year'],['Savings horizon',10,'years']];
s.getRange('F7:H19').values=assumptions;s.getRange('F7:F19').format.columnWidth=25;s.getRange('H7:H19').format.columnWidth=24;s.getRange('G10').setNumberFormat('0%');s.getRange('G12').setNumberFormat('0%');
val(s,'F22','No induced demand after the added lane.');val(s,'F23','All SW trips travel 8 miles and eventually arrive.');
val(s,'F25','Q10 refers to itself. Use Q9 annual savings × 10 ÷ 8.');val(s,'F26','Retain trip fractions; round only demand/lane for lookup.');
section(s,30,'Appendix B existing conditions');head(s,31,['Highway','Lanes','AM flow (veh/h)','AM speed (mph)','PM flow (veh/h)','PM speed (mph)']);
s.getRange('A32:F34').values=[['NE',6,10800,52,10200,53],['SW',4,6000,35,6400,39],['NW',5,8000,54,9000,52]];s.getRange('B32:F34').setNumberFormat('#,##0');band(s,32,34,6);
section(s,37,'Appendix C shares of new peak direction trips');head(s,38,['Period','NE share','SW share','NW share']);s.getRange('A39:D40').values=[['AM inbound',.35,.40,.25],['PM outbound',.30,.45,.25]];s.getRange('B39:D40').setNumberFormat('0%');band(s,39,40,4);
title(a,'A2','Office development traffic impact');
val(a,'A3','Demand and total flow are vehicles/hour; speed is mph. Rounded demand/lane selects Appendix A.');
section(a,5,'Q4 Updated Appendix B after development');
head(a,6,['Highway','Period','Lanes','Existing demand','New trips','Total demand','Demand / lane','Rounded demand / lane','Flow / lane','Total flow','Speed','LOS']);
const records=[['NE','AM',32,39,2],['SW','AM',33,39,3],['NW','AM',34,39,4],['NE','PM',32,40,2],['SW','PM',33,40,3],['NW','PM',34,40,4]];
for(let i=0;i<6;i++){
 let r=7+i,e=21+i,t=33+i;const [corr,period,src,shareRow,shareCol]=records[i];
 row(a,r,[corr,period]);fx(a,'C'+r,`=C${e}`);fx(a,'D'+r,`=H${e}`);fx(a,'E'+r,`=D${t}`);fx(a,'F'+r,`=D${r}+E${r}`);fx(a,'G'+r,`=F${r}/C${r}`);fx(a,'H'+r,`=ROUND(G${r},-2)`);
 fx(a,'I'+r,`=VLOOKUP(H${r},Inputs!$A$7:$D$27,2,FALSE)`);fx(a,'J'+r,`=I${r}*C${r}`);fx(a,'K'+r,`=VLOOKUP(H${r},Inputs!$A$7:$D$27,3,FALSE)`);fx(a,'L'+r,`=VLOOKUP(H${r},Inputs!$A$7:$D$27,4,FALSE)`);
 row(a,e,[corr,period]);fx(a,'C'+e,`=Inputs!B${src}`);fx(a,'D'+e,`=Inputs!${period==='AM'?'C':'E'}${src}`);fx(a,'E'+e,`=D${e}/C${e}`);fx(a,'F'+e,`=Inputs!${period==='AM'?'D':'F'}${src}`);
 fx(a,'G'+e,`=INDEX(Inputs!$A$7:$A$27,MATCH(F${e},Inputs!$C$7:$C$27,0))`);fx(a,'H'+e,`=G${e}*C${e}`);fx(a,'I'+e,`=INDEX(Inputs!$D$7:$D$27,MATCH(F${e},Inputs!$C$7:$C$27,0))`);
 row(a,t,[corr,period]);fx(a,'C'+t,`=Inputs!${String.fromCharCode(64+shareCol)}${shareRow}`);fx(a,'D'+t,`=C${t}*$D$${period==='AM'?30:31}`);
 let ch=41+i;row(a,ch,[corr,period]);fx(a,'C'+ch,`=I${e}`);fx(a,'D'+ch,`=L${r}`);fx(a,'E'+ch,`=F${r}-H${e}`);fx(a,'F'+ch,`=J${r}-D${e}`);fx(a,'G'+ch,`=K${r}-F${e}`);fx(a,'H'+ch,`=IF(C${ch}<>D${ch},"LOS worsens",IF(G${ch}<0,"Slower, same LOS","No speed change"))`);
}
band(a,7,12,12);a.getRange('C7:D12').setNumberFormat('#,##0');a.getRange('H7:K12').setNumberFormat('#,##0');
val(a,'A14','Use actual total demand for trip and time calculations. Total flow uses the rounded Appendix A row.');
val(a,'A15','Rounding can make estimated flow slightly exceed unrounded demand on uncongested corridors.');
val(a,'A16','SW demand rises while flow and speed fall in both peaks, even though LOS remains F.');
section(a,19,'Q1 Existing conditions and inferred demand');
head(a,20,['Highway','Period','Lanes','Observed total flow','Flow / lane','Speed','Demand / lane','Total demand','LOS']);band(a,21,26,9);a.getRange('C21:H26').setNumberFormat('#,##0');
section(a,28,'Q2 and Q3 Trip generation and assignment');
head(a,29,['Period','All directions','Peak share','Peak direction trips']);row(a,30,['AM inbound']);row(a,31,['PM outbound']);
fx(a,'B30','=Inputs!G7/Inputs!G8*Inputs!G9');fx(a,'C30','=Inputs!G10');fx(a,'D30','=B30*C30');fx(a,'B31','=Inputs!G7/Inputs!G8*Inputs!G11');fx(a,'C31','=Inputs!G12');fx(a,'D31','=B31*C31');a.getRange('C30:C31').setNumberFormat('0%');
head(a,32,['Highway','Period','Corridor share','Assigned trips']);a.getRange('C33:C38').setNumberFormat('0%');band(a,33,38,4);
section(a,39,'Q5 Changes relative to existing conditions');head(a,40,['Highway','Period','Old LOS','New LOS','Demand change','Flow change','Speed change','Result']);band(a,41,46,8);a.getRange('F41:G46').setNumberFormat('+#,##0;-#,##0;0');a.getRange('H41:H46').format.wrapText=true;a.getRange('A41:L46').format.rowHeight=34;
a.getRange('A1:B46').format.columnWidth=11;a.getRange('H1:H46').format.columnWidth=18;a.getRange('L1:L46').format.columnWidth=8;
title(l,'A2','Southwest highway added lane');
val(l,'A3','Q7–Q10 compare post-development conditions with four lanes against five lanes. Demand is unchanged.');
section(l,5,'Travel time savings');
row(l,6,['Daily vehicle-hours saved']);fx(l,'D6','=SUM(J23:J24)');
row(l,7,['Annual travel time value']);fx(l,'D7','=D6*Inputs!G17*Inputs!G18');
row(l,8,['Ten-year savings']);fx(l,'D8','=D7*Inputs!G19');
row(l,9,['Break-even cost per lane-mile']);fx(l,'D9','=D8/(Inputs!G13*Inputs!G14)');
l.getRange('A6:A9').format.font.bold=true;l.getRange('D6:D9').format.font.bold=true;l.getRange('D7:D9').setNumberFormat('$#,##0');l.getRange('D6').setNumberFormat('#,##0.00');
section(l,12,'Q7 Speeds and total vehicle flows');head(l,13,['Period','Total demand','Total lanes','Demand / lane','Rounded demand / lane','Flow / lane','Total flow','Speed (mph)','LOS']);
for(let i=0;i<2;i++){let r=14+i,p=i===0?8:11;val(l,'A'+r,i===0?'AM':'PM');fx(l,'B'+r,`='Traffic analysis'!F${p}`);fx(l,'C'+r,`='Traffic analysis'!C${p}+Inputs!$G$14`);fx(l,'D'+r,`=B${r}/C${r}`);fx(l,'E'+r,`=ROUND(D${r},-2)`);fx(l,'F'+r,`=VLOOKUP(E${r},Inputs!$A$7:$D$27,2,FALSE)`);fx(l,'G'+r,`=F${r}*C${r}`);fx(l,'H'+r,`=VLOOKUP(E${r},Inputs!$A$7:$D$27,3,FALSE)`);fx(l,'I'+r,`=VLOOKUP(E${r},Inputs!$A$7:$D$27,4,FALSE)`);}
band(l,14,15,9);l.getRange('C14:C15').setNumberFormat('0');l.getRange('E14:H15').setNumberFormat('#,##0');
val(l,'A17','AM improves from 32 to 48 mph and 5,600 to 9,500 veh/h; PM from 35 to 50 mph and 6,000 to 10,000 veh/h.');
section(l,21,'Q8 Demand based travel time calculation');head(l,22,['Period','Demand (veh/h)','Peak hours / day','4-lane speed','5-lane speed','Before (hr/trip)','After (hr/trip)','Saved (hr/trip)','Daily vehicles','Saved (veh-hr/day)']);
for(let i=0;i<2;i++){let r=23+i,p=i===0?8:11,n=14+i;val(l,'A'+r,i===0?'AM':'PM');fx(l,'B'+r,`=B${n}`);fx(l,'C'+r,`=Inputs!G${15+i}`);fx(l,'D'+r,`='Traffic analysis'!K${p}`);fx(l,'E'+r,`=H${n}`);fx(l,'F'+r,`=Inputs!$G$13/D${r}`);fx(l,'G'+r,`=Inputs!$G$13/E${r}`);fx(l,'H'+r,`=F${r}-G${r}`);fx(l,'I'+r,`=B${r}*C${r}`);fx(l,'J'+r,`=I${r}*H${r}`);}
band(l,23,24,10);l.getRange('C23:E24').setNumberFormat('0');l.getRange('F23:H24').setNumberFormat('0.000000');
val(l,'A26','Time saved per trip = 8 / old speed − 8 / new speed. Daily savings = hourly demand × peak hours × time saved.');
val(l,'A27','The 10-year comparison is undiscounted, as specified. The project adds 8 lane-miles.');
l.getRange('A1:A29').format.columnWidth=14;l.getRange('D1:D29').format.columnWidth=18;l.getRange('E1:E29').format.columnWidth=18;
for(const [sh,rs] of [[a,['L7:L12','I21:I26','C41:D46']],[l,['I14:I15']],[s,['D7:D27']]])for(const r of rs)sh.getRange(r).format.horizontalAlignment='center';
s.getRange('F6:H6').format.horizontalAlignment='center';
// Boundary and dependency tests, restoring every input before final verification.
const expect=(got,want,label,tol=1e-7)=>{if(Math.abs(got-want)>tol)throw Error(`${label}: ${got} != ${want}`);};
wb.recalculate();expect(l.getRange('D6').values[0][0],3113.4666666666667,'Daily savings');expect(l.getRange('D7').values[0][0],32691400,'Annual savings',.01);expect(l.getRange('D9').values[0][0],40864250,'Lane-mile cost',.01);
val(s,'G7',0);expect(a.getRange('D30').values[0][0],0,'Zero office trips');expect(a.getRange('F8').values[0][0],10000,'Zero project demand');
val(s,'G7',800000);val(s,'G14',0);expect(l.getRange('J23').values[0][0],0,'No added lane AM saving');expect(l.getRange('J24').values[0][0],0,'No added lane PM saving');val(s,'G14',1);
wb.recalculate();
console.log((await wb.inspect({kind:'table',range:"'Traffic analysis'!A7:L12",include:'values,formulas',tableMaxRows:6,tableMaxCols:12,maxChars:5000})).ndjson);
console.log((await wb.inspect({kind:'table',range:"'SW added lane'!A6:D9",include:'values,formulas',tableMaxRows:4,tableMaxCols:4,maxChars:1500})).ndjson);
console.log((await wb.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!',options:{useRegex:true,maxResults:30},summary:'Formula error scan',maxChars:1500})).ndjson);
const result=await SpreadsheetFile.exportXlsx(wb);await result.save(out+'Traffic Impact Calculations.xlsx');
for(const [sh,range,name] of [[a,'A1:L16','traffic-results'],[a,'A19:I46','traffic-work'],[l,'A1:J27','lane'],[s,'A1:J27','inputs-a'],[s,'A30:F40','inputs-bc']]){
 const img=await wb.render({sheetName:sh.name,range,scale:1.5,format:'png'});await fs.writeFile(out+'work/'+name+'.png',new Uint8Array(await img.arrayBuffer()));
}
console.log('Exported and rendered workbook.');

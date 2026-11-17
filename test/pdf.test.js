import test from 'node:test';
import assert from 'node:assert/strict';
import {normalize,analyze,sha256} from '../src/analysis.js';
import {portable,exportFiles} from '../src/export.js';
const params={from:'2026-04-01',to:'2026-04-30',zone:'America/Sao_Paulo',confirmed:true,area:{lat:-23,lon:-46,radius:100,label:'Estabelecimento de demonstração'},maxAccuracy:100,maxGapMinutes:30,maxSpeed:100,positionPairing:'hybrid'};
export function sampleReport(){
 const semanticSegments=Array.from({length:30},(_,i)=>{const day=String(i+1).padStart(2,'0');return [8,14].map(hour=>({startTime:`2026-04-${day}T${String(hour).padStart(2,'0')}:09:05-03:00`,endTime:`2026-04-${day}T${String(hour+3).padStart(2,'0')}:46:13-03:00`,visit:{topCandidate:{placeLocation:{latLng:'-23,-46'}}}}));}).flat();
 return {...portable(analyze(normalize({semanticSegments}),params),{name:'exemplo.json',sha256:'a'.repeat(64)}),id:'teste-pdf'};
}
test('exportação contém PDF legível, JSON e manifesto com hashes dos bytes exatos, sem XLSX',async()=>{
 const report=sampleReport(),files=await exportFiles(report);
 assert.deepEqual(files.map(f=>f.name),['geoprova-teste-pdf.json','geoprova-teste-pdf.pdf','geoprova-teste-pdf-manifesto.json']);
 assert.equal(new TextDecoder().decode(files[1].bytes.slice(0,5)),'%PDF-');
 assert.deepEqual(JSON.parse(new TextDecoder().decode(files[0].bytes)),JSON.parse(JSON.stringify(report)));
 const manifest=JSON.parse(new TextDecoder().decode(files[2].bytes));
 for(const entry of manifest.files){const file=files.find(f=>f.name===entry.name);assert.equal(entry.size,file.bytes.length);assert.equal(entry.sha256,await sha256(file.bytes));const altered=file.bytes.slice();altered[0]^=1;assert.notEqual(entry.sha256,await sha256(altered));}
});

import {regular,bold} from './pdf-fonts.js';
import {jsPDF} from 'jspdf';
import {autoTable} from 'jspdf-autotable';
import {displayDate,displayMonth,displayTimestamp,dayAt,toMinute,ceilMinute} from './dates.js';
import {dailyRows,groupedReportRows,summaryRows,timePairCount,periodTotals,totalNotes,totalDurationText,radiusReviewNotice} from './presentation.js';
import {userAlertText} from './export.js';

// Text and tables remain searchable; no screenshot, upload or print dialog is needed.
export function reportPDF(report) {
  const doc=new jsPDF({orientation:'landscape',unit:'mm',format:'a4'});
  doc.addFileToVFS('OpenSans-Regular.ttf',regular);doc.addFont('OpenSans-Regular.ttf','OpenSans','normal');
  doc.addFileToVFS('OpenSans-Semibold.ttf',bold);doc.addFont('OpenSans-Semibold.ttf','OpenSans','bold');
  doc.setFont('OpenSans');
  doc.setProperties({title:'Geoprova - Cartão de permanência',author:'Geoprova'});
  const width=doc.internal.pageSize.getWidth(),height=doc.internal.pageSize.getHeight(),margin=12;
  let y=18;
  const clean=value=>String(value??'').replaceAll('→',' para ').replaceAll('≈','').replace(/[–—]/g,'-');
  function paragraph(text,size=9,bold=false) {
    if(bold&&size>=12&&y>height-50){doc.addPage();y=18;}
    doc.setFont('OpenSans',bold?'bold':'normal');doc.setFontSize(size);
    for(const line of doc.splitTextToSize(clean(text),width-margin*2)) {
      if(y>height-17){doc.addPage();y=18;}
      doc.text(line,margin,y);y+=size*.45;
    }
    y+=3;
  }
  const table=(head,body)=>{
    if(y>height-35){doc.addPage();y=18;}
    autoTable(doc,{startY:y,head:[head],body,margin:{top:14,right:margin,bottom:16,left:margin},
      styles:{font:'OpenSans',fontSize:8,cellPadding:2,overflow:'linebreak',textColor:[32,59,58]},
      headStyles:{fillColor:[22,60,58],textColor:255},alternateRowStyles:{fillColor:[245,248,246]},
      rowPageBreak:'avoid',horizontalPageBreak:true,horizontalPageBreakRepeat:[0,1]});
    y=doc.lastAutoTable.finalY+8;
  };
  paragraph('Geoprova - Cartão de permanência',17,true);
  paragraph(report.params.area.label||'Local selecionado',11,true);
  paragraph(`${displayDate(report.params.from)} a ${displayDate(report.params.to)} | ${report.params.zone}`);
  paragraph('Horários em horas e minutos. Os segundos são descartados de cada entrada e saída antes do cálculo das durações e totais.');
  paragraph('Presença no local não equivale a tempo de trabalho. Dias sem registros não comprovam ausência.');
  if(report.days.some(d=>d.approximateMode))paragraph('Horários aproximados. A visita da Linha do Tempo tem prioridade. Sem visita no dia, a primeira e a última posição interna são usadas como limites conservadores, supondo permanência entre elas enquanto não houver evidência de saída ou deslocamento.');
  const rows=summaryRows(report),days=dailyRows(report),pairs=timePairCount(report);
  const radiusNotice=radiusReviewNotice(rows);if(radiusNotice){doc.setTextColor(180,35,24);paragraph('* Horário sem par. '+radiusNotice,9,true);doc.setTextColor(32,59,58);}
  const hasVisits=report.params.positionPairing==='spatial'||days.some(d=>d.visits.length),hasPositions=days.some(d=>d.positionTimes.length);
  const zone=report.params.zone;
  const clock=(t,day)=>t==null?'-':dayAt(t,zone)===day?new Intl.DateTimeFormat('pt-BR',{timeZone:zone,hour:'2-digit',minute:'2-digit'}).format(t):displayTimestamp(t,zone).replace(/(\d{2}:\d{2}):\d{2}/,'$1');
  const mark=number=>rows[number-1]?.alerts.length?` *${number}`:'';
  const boundary=(interval,kind,day)=>{
    if(!interval.spatial)return clock(kind==='arrival'?interval.arrival:interval.departure,day);
    const window=kind==='arrival'?interval.arrivalWindow:interval.departureWindow;
    return window?`${clock(toMinute(window[0]),day)} - ${clock(ceilMinute(window[1]),day)} (estimada)`:'Não determinada';
  };
  const timeCell=(interval,kind,day)=>{const value=kind==='arrival'?interval.arrival:interval.departure;const unpaired=!interval.spatial&&value!=null&&(interval.arrival==null||interval.departure==null);const content=clean(boundary(interval,kind,day))+(unpaired?' *':'')+(kind==='arrival'?(interval.sourceRowNumbers||[interval.rowNumber]).map(mark).join(''):'');return unpaired?{content,styles:{textColor:[180,35,24],fontStyle:'bold'}}:content;};
  const head=['Data / semana','Dia da semana',...(hasVisits?Array.from({length:pairs},(_,i)=>[`Entrada ${i+1}`,`Saída ${i+1}`]).flat():[]),...(hasPositions?['Posições registradas']:[]),'Total do dia'];
  const body=groupedReportRows(report).map(row=>{
    if(row.type==='week')return [{content:`Total da semana - ${displayDate(row.from)} a ${displayDate(row.to)}\n${clean(totalNotes(row))}`,colSpan:head.length-1,styles:{fillColor:[229,239,231],fontStyle:'bold'}},{content:clean(totalDurationText(row)),styles:{fillColor:[229,239,231],fontStyle:'bold'}}];
    return [displayDate(row.day)+(row.intervals.length?'':mark(row.emptyRowNumber)),row.weekday,
      ...(hasVisits?Array.from({length:pairs},(_,i)=>{const interval=row.visits[i];return interval?[timeCell(interval,'arrival',row.day),timeCell(interval,'departure',row.day)]:['-','-'];}).flat():[]),
      ...(hasPositions?[row.positionTimes.map(t=>clock(t,row.day)).join(' / ')+row.positionRows.map(p=>mark(p.rowNumber)).join('')||'-']:[]),clean(totalDurationText(row))];
  });
  table(head,body);
  paragraph('Totais mensais',12,true);
  table(['Mês','Permanência total','Cobertura do total'],periodTotals(report).months.map(m=>[displayMonth(m.key),clean(totalDurationText(m)),clean(totalNotes(m)||'Intervalos apurados no período selecionado')]));
  const alerts=rows.map((row,i)=>({row,number:i+1})).filter(x=>x.row.alerts.length);
  if(alerts.length){paragraph('Alertas dos registros',12,true);for(const {row,number}of alerts)paragraph(`*${number} | ${displayDate(row.day)} - ${userAlertText(row.alerts)}`);}
  paragraph('Arquivo e integridade',12,true);
  paragraph(`Identificador do relatório: ${report.id}\nVersão: ${report.appVersion}\nArquivo de origem: ${report.acquisition?.name||'Não informado'}\nSHA-256 do arquivo de origem: ${report.acquisition?.sha256||'Não informado'}`,8);
  paragraph('O manifesto de integridade acompanha o PDF e o JSON e permite conferir se os arquivos permanecem idênticos aos exportados. Não comprova autoria ou veracidade dos dados. Este PDF não possui assinatura digital.',8);
  paragraph('Limitações da análise',12,true);
  for(const note of report.limitations||[])paragraph(note,8);
  const count=doc.getNumberOfPages();
  for(let page=1;page<=count;page++){doc.setPage(page);doc.setFont('OpenSans','normal');doc.setFontSize(8);doc.setTextColor(90);doc.text(`Geoprova | Página ${page} de ${count}`,margin,height-7);}
  return new Uint8Array(doc.output('arraybuffer'));
}

async function accountStatement() {
  setTitle('Account Statement','Generate the MISB PDF account statement from your transaction log');
  const root=document.getElementById('content');
  root.innerHTML=`<div class="grid2"><div class="panel"><h2>Upload transaction log</h2><p>Upload the Cofundr Excel log to update the ledger and prepare your statement.</p><form id="pdfStatementForm"><label class="datepick">Transaction log (.xlsx, up to 10 MB)<input id="pdfLedgerFile" type="file" accept=".xlsx"></label><label class="datepick mt">Include transactions through<input id="pdfCutoff" type="date" value="${DATA.summary.last_transaction_date||''}"><small>Leave blank to use the latest transaction date.</small></label><button id="pdfPreviewButton" class="primary mt" type="submit">Prepare statement</button></form><p id="pdfStatementError" role="alert" class="danger"></p></div><div class="panel"><h2>Same MISB statement format</h2><p><b>Investor 5490 · Amanahraya Trustees Berhad</b></p><p>Cofundr letterhead, expanded account summary and six-column transaction detail, using the supplied statement layout.</p><p>Each profit payout is shown as Gross Profit, Service Charge, SST where applicable, and Net Profit. Deposits and withdrawals use their approved Excel entries; duplicate requests are excluded and withdrawal fees remain separate.</p><p class="muted">Uploading here replaces the workspace transaction ledger. The previous workbook version remains saved in D1.</p></div></div><div id="pdfStatementPreview" class="panel mt"><p>Prepare the statement to review totals and download the PDF.</p></div>`;
  const form=root.querySelector('#pdfStatementForm'), file=root.querySelector('#pdfLedgerFile'), cutoff=root.querySelector('#pdfCutoff'), button=root.querySelector('#pdfPreviewButton'), error=root.querySelector('#pdfStatementError'), preview=root.querySelector('#pdfStatementPreview');
  let prepared=null;
  const statementRM=value=>{
    if(value==='') return '';
    const amount=Number(value);
    return amount<0?`(${RM(Math.abs(amount))})`:RM(amount);
  };
  function invalidate(){prepared=null;preview.innerHTML='<p>Prepare the statement again to apply your changes.</p>';}
  file.addEventListener('change',()=>{cutoff.value='';invalidate();});
  cutoff.addEventListener('change',invalidate);
  form.addEventListener('submit',async event=>{
    event.preventDefault();button.disabled=true;error.textContent='';
    try {
      let response;
      if(file.files[0]) {
        const body=new FormData();body.append('file',file.files[0]);if(cutoff.value)body.append('as_of',cutoff.value);
        response=await fetch('/api/account-statement',{method:'POST',body});
      } else response=await fetch('/api/account-statement?'+new URLSearchParams(cutoff.value?{as_of:cutoff.value}:{}));
      prepared=await response.json();cutoff.value=prepared.as_of;
      const labels={starting_balance:'Starting balance',ending_balance:'Ending balance',total_investment:'Total investment',principal_received:'Total principal received',total_gross_returns:'Total gross returns',service_fee:'Service fee',sst:'SST',nett_returns:'Total nett returns received'};
      preview.innerHTML=`<div class="panelhead"><div><h2>Statement through ${esc(dmy(prepared.as_of))}</h2><p>${prepared.row_count} transactions · ${prepared.page_count} pages</p></div><button id="downloadStatementPdf" class="primary">Download account statement PDF</button></div>${prepared.warnings.map(w=>`<p class="banner">${esc(w)}</p>`).join('')}<div class="kpis">${Object.entries(labels).map(([k,label])=>kpi(label,RM(prepared.summary[k]),'')).join('')}</div><div class="tablewrap mt"><table><thead><tr><th>Transaction Date</th><th>Transaction Description</th><th>Note ID</th><th>Previous Balance (RM)</th><th>Sum Involved (RM)</th><th>Current Balance (RM)</th></tr></thead><tbody>${prepared.rows.slice(-12).map(r=>`<tr><td>${esc(r.date)}</td><td>${esc(r.description)}</td><td>${esc(r.note)}</td><td>${statementRM(r.previous)}</td><td>${statementRM(r.amount)}</td><td>${statementRM(r.current)}</td></tr>`).join('')}</tbody></table></div><p class="muted">Showing the last 12 transactions. The PDF includes all ${prepared.row_count} statement rows.</p>`;
      preview.querySelector('#downloadStatementPdf').addEventListener('click',async event=>{
        const download=event.currentTarget;download.disabled=true;download.textContent='Generating PDF...';
        try {
          const response=await fetch('/api/export/account-statement.pdf?'+new URLSearchParams({as_of:prepared.as_of,version:prepared.version}));
          const blob=await response.blob(),url=URL.createObjectURL(blob),link=document.createElement('a');
          link.href=url;link.download=`MISB_Account_Statement_${prepared.as_of}.pdf`;link.click();setTimeout(()=>URL.revokeObjectURL(url),10000);
        } catch(failure){error.textContent=failure.message;}
        finally{download.disabled=false;download.textContent='Download account statement PDF';}
      });
      if(file.files[0]) { DATA=await (await fetch('/api/data')).json();file.value=''; }
    } catch(failure){error.textContent=failure.message;}
    finally{button.disabled=false;}
  });
}

async function accountStatement() {
  setTitle('Account Statement','Generate and download your MISB account statement as PDF or CSV');
  const root=document.getElementById('content');
  root.innerHTML=`<div class="grid2"><div class="panel"><h2>Prepare account statement</h2><p>Upload the latest Cofundr Excel log to replace the ledger, or a prepared monthly CSV to reproduce its statement without changing the ledger. Balance gaps are kept as warnings; the previous workbook stays in history.</p><form id="pdfStatementForm"><label class="datepick">Transaction log (.xlsx) or prepared monthly statement (.csv), up to 10 MB<input id="pdfLedgerFile" type="file" accept=".xlsx,.csv"></label><label class="datepick mt">Statement type<select id="pdfMode"><option value="monthly">Monthly / next statement period</option><option value="all">Full transaction history</option></select></label><label class="datepick mt">Period starts<input id="pdfStart" type="date"><small>Leave blank to continue after the last report, or after 2 September 2026 for the first monthly report.</small></label><label class="datepick mt">Include transactions through<input id="pdfCutoff" type="date" value="${DATA.summary.last_transaction_date||''}"><small>Leave blank to use the latest transaction date.</small></label><button id="pdfPreviewButton" class="primary mt" type="submit">Prepare statement</button></form><p id="pdfStatementError" role="alert" class="danger"></p></div><div class="panel"><h2>Monthly MISB statement format</h2><p><b>Investor 5490 &middot; Amanahraya Trustees Berhad</b></p><p>Cofundr letterhead, account summary and transaction detail in the September 2026 format. PDF and CSV include the investor details, summary and all statement rows.</p><p>Each profit payout is shown as Gross Profit, Service Charge, SST where applicable, and Net Profit. Deposits and withdrawals use their approved Excel entries; duplicate requests are excluded and withdrawal fees remain separate.</p><p class="muted">Uploading here replaces the workspace transaction ledger after quality and regression checks. The previous workbook version remains saved in D1.</p></div></div><div id="pdfStatementPreview" class="panel mt"><p>Prepare the statement to review totals and download the PDF or CSV.</p></div><div id="statementHistory" class="panel mt"><div class="empty">Loading statement history&hellip;</div></div>`;
  const form=root.querySelector('#pdfStatementForm'), file=root.querySelector('#pdfLedgerFile'), cutoff=root.querySelector('#pdfCutoff'), button=root.querySelector('#pdfPreviewButton'), error=root.querySelector('#pdfStatementError'), preview=root.querySelector('#pdfStatementPreview');
  const mode=root.querySelector('#pdfMode'),start=root.querySelector('#pdfStart');
  let prepared=null,historyRuns=[],referenceFile=null;
  const history=root.querySelector('#statementHistory');
  async function loadHistory(){
    try{
      const response=await fetch('/api/account-statement-runs?limit=12'),runs=await response.json();historyRuns=runs;
      const verificationStatus=run=>run.last_verified_at?`${run.last_verification_matches?'<span class="pill normal">Record matches</span>':`<span class="pill overdue">Review ${esc((run.last_verification_mismatches||[]).join(', '))}</span>`}<br><span class="muted">${esc(run.last_verified_at)}</span>`:'<span class="muted">Not checked</span>';
      history.innerHTML=`<div class="panelhead"><div><h2>Generated statement history</h2><p>Each PDF remains linked to its integrity-checked transaction workbook. Verify recorded totals, download the source, or recreate the PDF without changing the live ledger.</p></div><a class="primary" href="/api/account-statement-runs.csv">Export statement register</a></div>${runs.length?`<div class="tablewrap compact"><table><thead><tr><th>Generated</th><th>Statement through</th><th class="num">Rows</th><th class="num">Pages</th><th class="num">Opening</th><th class="num">Closing</th><th class="num">Gross returns</th><th>Source workbook</th><th>Verification</th><th>Actions</th></tr></thead><tbody>${runs.map(run=>`<tr><td>${esc(run.created_at)}</td><td><b>${run.start_date?`${dmy(run.start_date)} to `:''}${dmy(run.as_of)}</b></td><td class="num">${run.row_count}</td><td class="num">${run.page_count}</td><td class="num">${RM(run.opening_balance)}</td><td class="num">${RM(run.closing_balance)}</td><td class="num">${RM(run.gross_returns)}</td><td><b>${esc(run.source_filename||'Transaction ledger')}</b><br>${run.source_available?`<span class="pill normal">Source verified</span> <span class="muted">SHA-256 ${esc(run.source_sha256.slice(0,12))}&hellip; &middot; ${(Number(run.source_byte_size)/1024).toFixed(1)} KB</span>`:'<span class="pill danger">Source unavailable</span>'}</td><td><span class="statement-verification-status">${verificationStatus(run)}</span><br>${run.source_available?`<button class="linkbtn verify-statement" data-id="${run.id}">${run.last_verified_at?'Reverify':'Verify record'}</button>`:''}</td><td><div class="version-actions">${run.source_available?`<a class="linkbtn" href="/api/source-versions/transactions/download?object_key=${encodeURIComponent(run.source_object_key)}">Source Excel</a><a class="linkbtn" href="/api/account-statement-runs/${run.id}/pdf">Recreate PDF</a>`:'<span class="muted">Cannot recreate</span>'}</div></td></tr>`).join('')}</tbody></table></div>`:'<p class="muted">No account-statement PDFs have been generated since history tracking was enabled.</p>'}`;
      history.querySelectorAll('.verify-statement').forEach(button=>button.addEventListener('click',async()=>{
        button.disabled=true;button.textContent='Verifying…';
        try{
          const response=await fetch(`/api/account-statement-runs/${button.dataset.id}/verify`,{method:'POST'}),result=await response.json();
          if(!response.ok)throw new Error(result.error||'Could not verify this statement.');
          button.closest('td').querySelector('.statement-verification-status').innerHTML=`${result.matches?'<span class="pill normal">Record matches</span>':'<span class="pill overdue">Review mismatch</span>'}<br><span class="muted">${esc(result.verified_at)}</span>`;
          button.disabled=false;button.textContent='Reverify';
        }catch(failure){button.disabled=false;button.textContent='Verify record';alert(failure.message)}
      }));
    }catch(failure){history.innerHTML='<p class="danger">Could not load statement history.</p>'}
  }
  const statementRM=value=>{
    if(value==='') return '';
    const amount=Number(value);
    return amount<0?`(${RM(Math.abs(amount))})`:RM(amount);
  };
  function invalidate(){prepared=null;referenceFile=null;preview.innerHTML='<p>Prepare the statement again to apply your changes.</p>';}
  file.addEventListener('change',()=>{cutoff.value='';const isReference=file.files[0]?.name.toLowerCase().endsWith('.csv');mode.disabled=Boolean(isReference);cutoff.disabled=Boolean(isReference);start.disabled=Boolean(isReference)||mode.value!=='monthly';invalidate();});
  cutoff.addEventListener('change',invalidate);
  start.addEventListener('change',invalidate);
  mode.addEventListener('change',()=>{start.disabled=mode.value!=='monthly';invalidate();});
  form.addEventListener('submit',async event=>{
    event.preventDefault();button.disabled=true;error.textContent='';
    try {
      let response;
      const params={mode:mode.value};if(cutoff.value)params.as_of=cutoff.value;if(mode.value==='monthly'&&start.value)params.start_date=start.value;
      if(file.files[0]) {
        const body=new FormData();body.append('file',file.files[0]);body.append('overwrite','1');body.append('allow_regression','1');Object.entries(params).forEach(([key,value])=>body.append(key,value));
        response=await fetch('/api/account-statement',{method:'POST',body});
      } else response=await fetch('/api/account-statement?'+new URLSearchParams(params));
      referenceFile=file.files[0]?.name.toLowerCase().endsWith('.csv')?file.files[0]:null;
      prepared=await response.json();if(!response.ok){const details=(prepared.issues||[]).filter(issue=>issue.level==='error').map(issue=>`${issue.message} (${issue.count})`);const reasons=prepared.comparison?.regression_reasons||[];throw new Error([prepared.error||'Could not prepare statement.',...details,...reasons].join(' '));}cutoff.value=prepared.as_of;start.value=prepared.start_date||'';
      const labels={starting_balance:'Starting balance',ending_balance:'Ending balance',total_investment:'Total investment',principal_received:'Total principal received',total_gross_returns:'Total gross returns',service_fee:'Service fee',sst:'SST',nett_returns:'Total nett returns received'};
      const recorded=historyRuns.some(run=>run.as_of===prepared.as_of&&(run.start_date||'')===(prepared.start_date||'')&&run.source_object_key===prepared.version);
      preview.innerHTML=`<div class="panelhead"><div><h2>Statement ${prepared.start_date?`${esc(dmy(prepared.start_date))} to `:'through '}${esc(dmy(prepared.as_of))}</h2><p>${esc(prepared.opening_source)} &middot; ${prepared.row_count} transactions &middot; ${prepared.page_count} pages ${referenceFile?'<span class="pill normal">Prepared CSV</span>':recorded?'<span class="pill normal">Already recorded</span>':'<span class="pill">New statement</span>'}</p></div><div class="version-actions"><button id="downloadStatementPdf" class="primary">Download account statement PDF</button><button id="downloadStatementCsv" class="primary">Download CSV</button></div></div><p id="pdfStatementNotice" class="ok"></p>${prepared.warnings.map(w=>`<p class="banner">${esc(w)}</p>`).join('')}<div class="kpis">${Object.entries(labels).map(([k,label])=>kpi(label,RM(prepared.summary[k]),'')).join('')}</div><div class="tablewrap mt"><table><thead><tr><th>Transaction Date</th><th>Note ID</th><th>Transaction Description</th><th>Previous Balance (RM)</th><th>Sum Involved (RM)</th><th>Current Balance (RM)</th></tr></thead><tbody>${prepared.rows.slice(-12).map(r=>`<tr><td>${esc(r.date)}</td><td>${esc(r.note)}</td><td>${esc(r.description)}</td><td>${statementRM(r.previous)}</td><td>${statementRM(r.amount)}</td><td>${statementRM(r.current)}</td></tr>`).join('')}</tbody></table></div><p class="muted">Showing the last 12 transactions. The PDF and CSV include all ${prepared.row_count} statement rows.</p>`;
      preview.querySelectorAll('#downloadStatementPdf, #downloadStatementCsv').forEach(downloadButton=>downloadButton.addEventListener('click',async event=>{
        const format=downloadButton.id==='downloadStatementCsv'?'csv':'pdf',label=downloadButton.textContent;
        const download=event.currentTarget;download.disabled=true;download.textContent=`Generating ${format.toUpperCase()}...`;
        try {
          const body=new FormData();if(referenceFile)body.append('file',referenceFile);
          const response=await fetch(`/api/export/account-statement.${format}?`+new URLSearchParams({as_of:prepared.as_of,version:prepared.version,mode:prepared.mode||'all',start_date:prepared.start_date||''}),referenceFile?{method:'POST',body}:{});
          if(!response.ok){let message='Could not generate the account statement.';try{message=(await response.json()).error||message}catch(parseError){}throw new Error(message)}
          const isNew=response.headers.get('X-Statement-History-New')==='1';
          const blob=await response.blob(),url=URL.createObjectURL(blob),link=document.createElement('a');
          link.href=url;link.download=`MISB_Account_Statement_${prepared.start_date?prepared.start_date+'_to_':''}${prepared.as_of}.${format}`;link.click();setTimeout(()=>URL.revokeObjectURL(url),10000);
          if(format==='pdf'&&!referenceFile)await loadHistory();
          preview.querySelector('#pdfStatementNotice').textContent=referenceFile?'Prepared statement downloaded.':format==='csv'?'CSV downloaded with all statement transactions.':isNew?'Statement generated and recorded in history.':'Downloaded the existing recorded statement; no duplicate history row was added.';
        } catch(failure){error.textContent=failure.message;}
        finally{download.disabled=false;download.textContent=label;}
      }));
      if(file.files[0]&&!referenceFile) { DATA=await (await fetch('/api/data')).json();file.value=''; }
    } catch(failure){error.textContent=failure.message;}
    finally{button.disabled=false;}
  });
  await loadHistory();
}

async function allocationEmailMonitor(){
  setTitle('Allocation Email Monitor','Track allocation emails and disbursement emails against the transaction log');
  const root=document.getElementById('content');
  root.innerHTML='<div class="panel"><div class="empty">Checking ledger allocations…</div></div>';
  try{
    const response=await fetch('/api/allocation-email-tracking'),data=await response.json();
    if(!response.ok)throw new Error(data.error||'Could not load the email monitor.');
    if(view!=='email-tracking')return;
    const summary=data.summary;
    root.innerHTML=`<div class="kpis">${kpi('Allocated notes',summary.total,'Successful Investment Committed transactions')}${kpi('Emails complete',summary.emails_complete,'Both email dates recorded and allocations reviewed')}${kpi('Emails to check',summary.pending,'Missing dates or changed allocations')}${kpi('Missing disbursement date',summary.missing_disbursal,'Complete the simulation or campaign entry')}</div>
      <div class="panel"><div class="panelhead"><div><h2>Allocation email register</h2><p>One row per note, including every successful allocation in the current transaction log. Email dates start from the imported simulation. Enter the actual sent dates and save to confirm they cover all allocations listed for that note. Saved dates appear in the generated simulation report.</p><p class="muted">This register records your confirmation; it does not check your mailbox. Disbursement dates come from the simulation or campaign entry. Clear an email date and save to mark it pending again.</p></div></div>
      ${data.unmatched.length?`<div class="banner"><b>${data.unmatched.length} allocation transaction(s) have no note reference.</b> Assign them in Transaction Matching before confirming all emails are complete.</div>`:''}
      <div class="audit-filters"><input id="emailMonitorSearch" type="search" placeholder="Search note or issuer" aria-label="Search allocation emails"><select id="emailMonitorFilter" aria-label="Filter allocation emails"><option value="pending">Needs attention</option><option value="all">All allocations</option><option value="complete">Emails complete</option></select><span id="emailMonitorCount"></span></div>
      <p id="emailMonitorMessage" role="status" class="muted"></p><div class="tablewrap"><table><thead><tr><th>Note / Issuer</th><th>Ledger allocation dates</th><th>Allocated amount</th><th>Allocation email sent</th><th>Disbursement email sent</th><th>Disbursement date</th><th>Status</th><th></th></tr></thead><tbody id="emailMonitorRows"></tbody></table></div></div>`;
    const search=root.querySelector('#emailMonitorSearch'),filter=root.querySelector('#emailMonitorFilter'),body=root.querySelector('#emailMonitorRows'),message=root.querySelector('#emailMonitorMessage');
    function draw(){
      const query=search.value.trim().toLowerCase(),rows=data.rows.filter(row=>`${row.loan_code} ${row.issuer} ${row.note_name}`.toLowerCase().includes(query)&&(filter.value==='all'||(filter.value==='complete'?row.emails_complete:row.pending.length)));
      root.querySelector('#emailMonitorCount').textContent=`${rows.length} of ${data.rows.length} notes`;
      body.innerHTML=rows.map(row=>`<tr data-code="${esc(row.loan_code)}"><td><b>${esc(row.loan_code)}</b><br>${esc(row.issuer||row.note_name||'Campaign details missing')}<br><small class="muted">${esc(row.source)}</small></td><td>${row.allocation_dates.map(dmy).join('<br>')||'Missing date'}<br><small class="muted">${row.transaction_count} transaction(s)</small></td><td>${RM(row.amount)}</td><td><input class="email-date" type="date" name="allocation_email_date" aria-label="Allocation email sent for ${esc(row.loan_code)}" value="${esc(row.allocation_email_date)}" max="${DATA.today}"></td><td><input class="email-date" type="date" name="disbursement_email_date" aria-label="Disbursement email sent for ${esc(row.loan_code)}" value="${esc(row.disbursement_email_date)}" max="${DATA.today}"></td><td>${row.disbursal_date?dmy(row.disbursal_date):'<span class="danger">Missing</span>'}</td><td>${row.pending.length?row.pending.map(text=>`<div class="muted">${esc(text)}</div>`).join(''):'<span class="pill normal">Complete</span>'}</td><td><button type="button" class="primary save-email-dates">Save / confirm</button><span class="email-save-status" role="status"></span></td></tr>`).join('')||`<tr><td colspan="8" class="empty">${data.rows.length?'No allocations match this filter.':'No successful allocations found. Import the transaction log to begin.'}</td></tr>`;
    }
    body.addEventListener('input',event=>{
      const tr=event.target.closest('tr'),row=data.rows.find(item=>item.loan_code===tr.dataset.code);
      row[event.target.name]=event.target.value;
      tr.querySelector('.email-save-status').textContent='Unsaved';
    });
    body.addEventListener('click',async event=>{
      const button=event.target.closest('.save-email-dates');if(!button)return;
      const tr=button.closest('tr'),row=data.rows.find(item=>item.loan_code===tr.dataset.code);
      if(!Array.from(tr.querySelectorAll('input')).every(input=>input.reportValidity()))return;
      button.disabled=true;message.textContent='';
      try{
        const response=await fetch('/api/allocation-email-tracking/'+encodeURIComponent(row.loan_code),{method:'PUT',headers:{'Content-Type':'application/json'},body:JSON.stringify({allocation_email_date:row.allocation_email_date,disbursement_email_date:row.disbursement_email_date,ledger_signature:row.ledger_signature})});
        const result=await response.json();if(!response.ok)throw new Error(result.error||'Could not save email dates.');
        const refreshed=await fetch('/api/allocation-email-tracking');const current=await refreshed.json();if(!refreshed.ok)throw new Error(current.error||'Saved, but could not refresh the register.');
        const updated=current.rows.find(item=>item.loan_code===row.loan_code);Object.assign(row,updated);
        root.querySelector('.kpis').innerHTML=`${kpi('Allocated notes',current.summary.total,'Successful Investment Committed transactions')}${kpi('Emails complete',current.summary.emails_complete,'Both email dates recorded and allocations reviewed')}${kpi('Emails to check',current.summary.pending,'Missing dates or changed allocations')}${kpi('Missing disbursement date',current.summary.missing_disbursal,'Complete the simulation or campaign entry')}`;
        message.textContent=`${row.loan_code}: email dates saved.`;draw();
      }catch(error){message.textContent=error.message;}finally{button.disabled=false;}
    });
    search.addEventListener('input',draw);filter.addEventListener('change',draw);draw();
  }catch(error){if(view==='email-tracking')root.innerHTML=`<div class="panel"><p class="danger" role="alert">${esc(error.message)}</p></div>`;}
}

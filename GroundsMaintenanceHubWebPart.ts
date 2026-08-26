
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { SharePointService, ICustomer, ISite, ICrew, IJob } from './SharePointService';

export interface IGroundsMaintenanceHubWebPartProps {}

export default class GroundsMaintenanceHubWebPart extends BaseClientSideWebPart<IGroundsMaintenanceHubWebPartProps> {
  private service!: SharePointService;
  private customers: ICustomer[] = [];
  private sites: ISite[] = [];
  private crews: ICrew[] = [];
  private jobs: IJob[] = [];
  private currentView = 'dashboard';
  private scheduleStart = new Date();

  public async onInit(): Promise<void> {
    this.service = new SharePointService(this.context);
    await this.loadData();
  }

  private async loadData(): Promise<void> {
    [this.customers, this.sites, this.crews, this.jobs] = await Promise.all([
      this.service.getCustomers(),
      this.service.getSites(),
      this.service.getCrews(),
      this.service.getJobs()
    ]);
  }

  public render(): void {
    this.domElement.innerHTML = `
      <div class="gmhub">
        <div class="gmhub-header">
          <div>
            <div class="gmhub-title">Grounds Maintenance Hub</div>
            <div class="gmhub-subtitle">Schedule • Complete • Invoice • Report</div>
          </div>
          <div class="gmhub-save">✓ Saved in SharePoint</div>
        </div>
        <div class="gmhub-tabs">
          ${['dashboard','schedule','jobs','customers'].map(v => `
            <button data-view="${v}" class="${this.currentView===v?'active':''}">
              ${v.charAt(0).toUpperCase()+v.slice(1)}
            </button>`).join('')}
        </div>
        <div id="gmhub-body">${this.renderView()}</div>
      </div>
      ${this.styles()}
    `;
    this.bindEvents();
  }

  private renderView(): string {
    switch(this.currentView) {
      case 'schedule': return this.renderSchedule();
      case 'jobs': return this.renderJobs();
      case 'customers': return this.renderCustomers();
      default: return this.renderDashboard();
    }
  }

  private money(v:number): string {
    return new Intl.NumberFormat('en-AU',{style:'currency',currency:'AUD',minimumFractionDigits:2,maximumFractionDigits:2}).format(v||0);
  }

  private renderDashboard(): string {
    const active = this.jobs.filter(j => j.status !== 'Cancelled');
    const scheduled = active.filter(j => j.status === 'Scheduled').length;
    const completed = active.filter(j => j.status === 'Completed').length;
    const invoiced = active.filter(j => j.status === 'Invoiced').length;

    // Monthly flat contracts count once; per-visit revenue comes from jobs.
    const contractRevenue = this.customers
      .filter(c => c.billingType === 'Monthly Flat')
      .reduce((a,c) => a + Number(c.monthlyContractValue || c.defaultRate || 0),0);

    const variableRevenue = active
      .filter(j => j.status === 'Completed' || j.status === 'Invoiced')
      .filter(j => j.billingType !== 'Monthly Flat' && j.billingType !== 'Contract Visit')
      .reduce((a,j) => a + Number(j.invoiceValue || 0),0);

    const revenue = contractRevenue + variableRevenue;

    const byCustomer: Record<string,number> = {};
    this.customers.forEach(c => byCustomer[c.name] = c.billingType === 'Monthly Flat'
      ? Number(c.monthlyContractValue || c.defaultRate || 0) : 0);

    active.filter(j => j.status === 'Completed' || j.status === 'Invoiced').forEach(j => {
      if (j.billingType !== 'Monthly Flat' && j.billingType !== 'Contract Visit') {
        byCustomer[j.customer] = (byCustomer[j.customer] || 0) + Number(j.invoiceValue || 0);
      }
    });

    const rows = Object.entries(byCustomer).filter(([,v])=>v>0).sort((a,b)=>b[1]-a[1]).slice(0,10);
    const max = Math.max(1,...rows.map(r=>r[1]));

    return `
      <h2>Dashboard</h2>
      <div class="kpis">
        ${this.kpi('Scheduled Visits', String(scheduled))}
        ${this.kpi('Completed Visits', String(completed))}
        ${this.kpi('Invoiced Jobs', String(invoiced))}
        ${this.kpi('Revenue This Month', this.money(revenue))}
      </div>
      <div class="panel">
        <h3>Revenue by Customer</h3>
        ${rows.length ? rows.map(([name,value]) => `
          <div class="chart-row">
            <div class="chart-label">${name}</div>
            <div class="chart-track"><div class="chart-fill" style="width:${Math.max(2,value/max*100)}%"></div></div>
            <div class="chart-value">${this.money(value)}</div>
          </div>`).join('') : '<div class="empty">No revenue data yet</div>'}
      </div>`;
  }

  private kpi(label:string,value:string): string {
    return `<div class="kpi"><div class="kpi-label">${label}</div><div class="kpi-value">${value}</div></div>`;
  }

  private iso(d:Date): string {
    return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
  }

  private renderSchedule(): string {
    const dates: Date[] = [];
    for(let i=0;i<14;i++){
      const d = new Date(this.scheduleStart);
      d.setDate(this.scheduleStart.getDate()+i);
      dates.push(d);
    }

    let grid = `<div class="cal-grid"><div class="cal-head">Crew</div>`;
    dates.forEach(d => {
      grid += `<div class="cal-head">${d.toLocaleDateString('en-AU',{weekday:'short',day:'2-digit',month:'short'})}</div>`;
    });

    this.crews.forEach(crew => {
      grid += `<div class="cal-crew">${crew.name}</div>`;
      dates.forEach(d => {
        const date = this.iso(d);
        const list = this.jobs.filter(j =>
          (j.scheduledDate || '').slice(0,10) === date &&
          j.crew === crew.name &&
          j.status !== 'Cancelled'
        );
        grid += `<div class="cal-cell">`;
        if (list.length) {
          list.forEach(j => {
            grid += `<div class="job-card">
              <b>${j.customer}</b>
              <span>${j.site}</span>
              <span>${j.jobDescription || j.title}</span>
              ${j.psiNumber ? `<span>PSI: ${j.psiNumber}</span>` : ''}
              <select data-job-status="${j.id}">
                ${['Scheduled','Completed','Invoiced','Cancelled'].map(s =>
                  `<option ${j.status===s?'selected':''}>${s}</option>`).join('')}
              </select>
            </div>`;
          });
        } else {
          grid += `<div class="empty">—</div>`;
        }
        grid += `</div>`;
      });
    });
    grid += `</div>`;

    return `
      <div class="toolbar">
        <div>
          <h2>14 Day Schedule</h2>
          <div class="muted">All active crews from GM Crews.</div>
        </div>
        <label>Start date <input id="scheduleStart" type="date" value="${this.iso(this.scheduleStart)}"></label>
        <button id="addJob">+ Schedule Job</button>
      </div>
      <div class="schedule-wrap">${grid}</div>`;
  }

  private renderJobs(): string {
    const sorted = [...this.jobs].sort((a,b)=>(a.scheduledDate||'').localeCompare(b.scheduledDate||''));
    return `
      <h2>Jobs</h2>
      <div class="panel table-panel">
        <table>
          <thead><tr>
            <th>Date</th><th>Customer</th><th>Site</th><th>Crew</th><th>Job</th>
            <th>Billing</th><th>Invoice Value</th><th>Cost</th><th>PSI</th><th>Status</th>
          </tr></thead>
          <tbody>
          ${sorted.map(j => `<tr>
            <td>${(j.scheduledDate||'').slice(0,10)}</td>
            <td>${j.customer}</td><td>${j.site}</td><td>${j.crew}</td>
            <td>${j.jobDescription || j.title}</td><td>${j.billingType||''}</td>
            <td>${this.money(Number(j.invoiceValue||0))}</td>
            <td>${this.money(Number(j.jobCost||0))}</td>
            <td>${j.psiNumber||''}</td>
            <td><select data-job-status="${j.id}">
              ${['Scheduled','Completed','Invoiced','Cancelled'].map(s =>
                `<option ${j.status===s?'selected':''}>${s}</option>`).join('')}
            </select></td>
          </tr>`).join('')}
          </tbody>
        </table>
      </div>`;
  }

  private renderCustomers(): string {
    return `
      <h2>Customers</h2>
      <div class="panel table-panel">
        <table>
          <thead><tr><th>Customer</th><th>Billing</th><th>Default Rate</th><th>Monthly Contract</th><th>PO</th><th>Sites</th></tr></thead>
          <tbody>${this.customers.map(c => {
            const siteNames = this.sites.filter(s => s.customerId === c.id || s.customerName === c.name).map(s=>s.name);
            return `<tr><td>${c.name}</td><td>${c.billingType||''}</td>
              <td>${this.money(Number(c.defaultRate||0))}</td>
              <td>${this.money(Number(c.monthlyContractValue||0))}</td>
              <td>${c.poNumber||''}</td><td>${siteNames.join(', ')}</td></tr>`;
          }).join('')}</tbody>
        </table>
      </div>`;
  }

  private async addJob(): Promise<void> {
    const customer = prompt('Customer name');
    if (!customer) return;
    const relevantSites = this.sites.filter(s => s.customerName === customer);
    const site = prompt(`Site (${relevantSites.map(s=>s.name).join(', ')})`);
    if (!site) return;
    const crew = prompt(`Crew (${this.crews.map(c=>c.name).join(', ')})`);
    if (!crew) return;
    const scheduledDate = prompt('Scheduled date (YYYY-MM-DD)', this.iso(this.scheduleStart));
    if (!scheduledDate) return;
    const desc = prompt('Job description','Grounds Maintenance') || 'Grounds Maintenance';

    const c = this.customers.find(x=>x.name===customer);
    const value = Number(c?.defaultRate || 0);
    const billing = c?.billingType || 'Per Visit';

    await this.service.createJob({
      customer,
      site,
      crew,
      scheduledDate,
      jobDescription: desc,
      billingType: billing,
      invoiceValue: billing === 'Monthly Flat' ? 0 : value,
      poNumber: c?.poNumber || '',
      status:'Scheduled'
    });

    await this.loadData();
    this.render();
  }

  private bindEvents(): void {
    this.domElement.querySelectorAll('[data-view]').forEach(el => {
      el.addEventListener('click', () => {
        this.currentView = (el as HTMLElement).getAttribute('data-view') || 'dashboard';
        this.render();
      });
    });

    this.domElement.querySelectorAll('[data-job-status]').forEach(el => {
      el.addEventListener('change', async (ev:any) => {
        const id = Number((ev.target as HTMLElement).getAttribute('data-job-status'));
        const status = ev.target.value;
        try {
          await this.service.updateJobStatus(id,status);
          await this.loadData();
          this.render();
        } catch (e:any) {
          alert(e.message || 'Could not save status.');
        }
      });
    });

    const dateEl = this.domElement.querySelector('#scheduleStart') as HTMLInputElement | null;
    if (dateEl) {
      dateEl.addEventListener('change', () => {
        this.scheduleStart = new Date(dateEl.value+'T00:00:00');
        this.render();
      });
    }

    const add = this.domElement.querySelector('#addJob');
    if (add) add.addEventListener('click', () => this.addJob());
  }

  private styles(): string {
    return `<style>
      .gmhub{font-family:Segoe UI,Arial,sans-serif;background:#f5f7f9;color:#243447;min-height:600px}
      .gmhub-header{background:#58a447;color:#fff;padding:18px 22px;display:flex;justify-content:space-between;align-items:center}
      .gmhub-title{font-size:24px;font-weight:800}.gmhub-subtitle{font-size:12px;opacity:.9}.gmhub-save{font-size:12px;font-weight:700}
      .gmhub-tabs{background:#fff;border-bottom:1px solid #dfe5ea;padding:8px 14px;display:flex;gap:8px}
      .gmhub-tabs button,.toolbar button{border:0;background:#3e75b6;color:#fff;padding:9px 14px;border-radius:7px;font-weight:700;cursor:pointer}
      .gmhub-tabs button.active{background:#233650}
      #gmhub-body{padding:20px}.muted{color:#6b7685;font-size:12px}.toolbar{display:flex;justify-content:space-between;gap:12px;align-items:end;margin-bottom:14px}
      .toolbar input{padding:8px;border:1px solid #ccd5dd;border-radius:6px}
      .kpis{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin:15px 0}
      .kpi,.panel{background:#fff;border:1px solid #dfe5ea;border-radius:10px;padding:16px;box-shadow:0 2px 7px #0000000a}
      .kpi-label{font-size:11px;text-transform:uppercase;color:#6b7685;font-weight:700}.kpi-value{font-size:28px;font-weight:800;color:#233650;margin-top:7px}
      .chart-row{display:grid;grid-template-columns:minmax(150px,1.3fr) 3fr auto;gap:10px;align-items:center;margin:11px 0}
      .chart-label{font-size:12px;font-weight:700}.chart-track{height:16px;background:#e7edf2;border-radius:999px;overflow:hidden}
      .chart-fill{height:100%;background:#58a447;border-radius:999px}.chart-value{font-size:12px;font-weight:800}
      .schedule-wrap{overflow:auto;background:#fff;border:1px solid #dfe5ea;border-radius:10px}.cal-grid{display:grid;grid-template-columns:150px repeat(14,125px);min-width:1900px}
      .cal-head{background:#233650;color:#fff;font-size:11px;font-weight:800;padding:9px;text-align:center;border-right:1px solid #42536a}
      .cal-crew{background:#eef3f6;font-size:12px;font-weight:800;padding:9px;border-right:1px solid #dfe5ea;border-bottom:1px solid #dfe5ea;display:flex;align-items:center;justify-content:center}
      .cal-cell{min-height:110px;padding:5px;border-right:1px solid #dfe5ea;border-bottom:1px solid #dfe5ea}
      .job-card{padding:6px;margin-bottom:5px;border:1px solid #dfe5ea;border-left:4px solid #3e75b6;border-radius:6px;background:#fff}
      .job-card b,.job-card span{display:block}.job-card span{font-size:10px;color:#6b7685;margin-top:2px}.job-card select{width:100%;font-size:10px;margin-top:5px}
      .empty{color:#a0a8b2;text-align:center;padding:15px 2px}.table-panel{overflow:auto;padding:0}
      table{width:100%;border-collapse:collapse}th{background:#233650;color:#fff;text-align:left;padding:9px;font-size:11px}td{padding:9px;border-bottom:1px solid #dfe5ea;font-size:12px}
      td select{padding:5px;border:1px solid #cfd7de;border-radius:5px}
      @media(max-width:900px){.kpis{grid-template-columns:1fr 1fr}}
    </style>`;
  }
}

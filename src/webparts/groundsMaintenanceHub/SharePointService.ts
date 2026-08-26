
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface ICustomer {
  id: number;
  name: string;
  customerNumber?: string;
  mainContact?: string;
  email?: string;
  phone?: string;
  active: boolean;
  billingType?: string;
  defaultRate?: number;
  poNumber?: string;
  monthlyContractValue?: number;
}

export interface ISite {
  id: number;
  name: string;
  customerId?: number;
  customerName?: string;
  address?: string;
  siteContact?: string;
  phone?: string;
  active: boolean;
}

export interface ICrew {
  id: number;
  name: string;
  active: boolean;
}

export interface IJob {
  id: number;
  title: string;
  customer: string;
  site: string;
  scheduledDate: string;
  jobDescription: string;
  crew: string;
  billingType?: string;
  invoiceValue?: number;
  jobCost?: number;
  poNumber?: string;
  status: string;
  completionNotes?: string;
  completedDate?: string;
  salesOrderNumber?: string;
  invoicedDate?: string;
  readyForInvoiceDate?: string;
  psiNumber?: string;
}

type FieldMap = Record<string, string>;

export class SharePointService {
  private readonly siteUrl: string;
  private fieldCache: Record<string, FieldMap> = {};

  constructor(private readonly context: WebPartContext) {
    this.siteUrl = 'https://wavind.sharepoint.com/sites/Waverley';
  }

  private async json(url: string, options?: any): Promise<any> {
    const response: SPHttpClientResponse = await this.context.spHttpClient.get(
      url,
      SPHttpClient.configurations.v1,
      options
    );
    if (!response.ok) {
      throw new Error(`SharePoint request failed: ${response.status} ${response.statusText}`);
    }
    return response.json();
  }

  private async fieldMap(listTitle: string): Promise<FieldMap> {
    if (this.fieldCache[listTitle]) return this.fieldCache[listTitle];

    const url =
      `${this.siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')` +
      `/fields?$select=Title,InternalName,Hidden`;

    const data = await this.json(url);
    const map: FieldMap = {};
    (data.value || []).forEach((f: any) => {
      if (!f.Hidden) map[f.Title] = f.InternalName;
    });
    this.fieldCache[listTitle] = map;
    return map;
  }

  private async listItems(listTitle: string, selectTitles: string[]): Promise<any[]> {
    const map = await this.fieldMap(listTitle);
    const internal = selectTitles
      .map(t => map[t])
      .filter(Boolean);

    const select = ['Id', ...internal].join(',');
    const url =
      `${this.siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')` +
      `/items?$top=5000&$select=${encodeURIComponent(select)}`;

    const data = await this.json(url);
    return data.value || [];
  }

  public async getCustomers(): Promise<ICustomer[]> {
    const list = 'GM Customers';
    const map = await this.fieldMap(list);
    const rows = await this.listItems(list, [
      'Title','Customer Number','Main Contact','Email','Phone','Active',
      'Billing Type','Default Rate','PO Number','Monthly Contract Value'
    ]);

    const v = (r:any, title:string) => r[map[title]];
    return rows.map((r:any) => ({
      id: r.Id,
      name: v(r,'Title') || '',
      customerNumber: v(r,'Customer Number') || '',
      mainContact: v(r,'Main Contact') || '',
      email: v(r,'Email') || '',
      phone: v(r,'Phone') || '',
      active: v(r,'Active') !== false,
      billingType: v(r,'Billing Type') || '',
      defaultRate: Number(v(r,'Default Rate') || 0),
      poNumber: v(r,'PO Number') || '',
      monthlyContractValue: Number(v(r,'Monthly Contract Value') || 0)
    })).filter((x: ISite) => x.active);
  }

  public async getSites(): Promise<ISite[]> {
    const list = 'GM Sites';
    const map = await this.fieldMap(list);

    // Lookup fields need Id and display value. Resolve Customer's internal name dynamically.
    const customerInternal = map['Customer'];
    const select = [
      'Id',
      map['Title'],
      map['Address'],
      map['Site Contact'],
      map['Phone'],
      map['Active'],
      `${customerInternal}/Id`,
      `${customerInternal}/Title`
    ].filter(Boolean).join(',');

    const expand = customerInternal ? `&$expand=${customerInternal}` : '';
    const url =
      `${this.siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(list)}')` +
      `/items?$top=5000&$select=${encodeURIComponent(select)}${expand}`;

    const data = await this.json(url);
    const rows = data.value || [];
    const v = (r:any, title:string) => r[map[title]];

    return rows.map((r:any) => ({
      id: r.Id,
      name: v(r,'Title') || '',
      customerId: r[customerInternal]?.Id,
      customerName: r[customerInternal]?.Title || '',
      address: v(r,'Address') || '',
      siteContact: v(r,'Site Contact') || '',
      phone: v(r,'Phone') || '',
      active: v(r,'Active') !== false
    })).filter((x: ISite) => x.active);
  }

  public async getCrews(): Promise<ICrew[]> {
    const list = 'GM Crews';
    const map = await this.fieldMap(list);
    const rows = await this.listItems(list, ['Title','Active']);
    const v = (r:any, title:string) => r[map[title]];

    return rows.map((r:any) => ({
      id: r.Id,
      name: v(r,'Title') || '',
      active: v(r,'Active') !== false
})).filter((x: ICrew) => x.active);
  }

  public async getJobs(): Promise<IJob[]> {
    const list = 'GM Jobs';
    const map = await this.fieldMap(list);
    const titles = [
      'Title','Customer','Site','Scheduled Date','Job Description','Crew',
      'Billing Type','Invoice Value','Job Cost','PO Number','Status',
      'Completion Notes','Completed Date','Sales Order Number','Invoiced Date',
      'Ready for Invoice Date','PSI Number'
    ];
    const rows = await this.listItems(list, titles);
    const v = (r:any, title:string) => r[map[title]];

    return rows.map((r:any) => ({
      id: r.Id,
      title: v(r,'Title') || '',
      customer: v(r,'Customer') || '',
      site: v(r,'Site') || '',
      scheduledDate: v(r,'Scheduled Date') || '',
      jobDescription: v(r,'Job Description') || '',
      crew: v(r,'Crew') || '',
      billingType: v(r,'Billing Type') || '',
      invoiceValue: Number(v(r,'Invoice Value') || 0),
      jobCost: Number(v(r,'Job Cost') || 0),
      poNumber: v(r,'PO Number') || '',
      status: v(r,'Status') || 'Scheduled',
      completionNotes: v(r,'Completion Notes') || '',
      completedDate: v(r,'Completed Date') || '',
      salesOrderNumber: v(r,'Sales Order Number') || '',
      invoicedDate: v(r,'Invoiced Date') || '',
      readyForInvoiceDate: v(r,'Ready for Invoice Date') || '',
      psiNumber: v(r,'PSI Number') || ''
    }));
  }

  private async digest(): Promise<string> {
    const response = await this.context.spHttpClient.post(
      `${this.siteUrl}/_api/contextinfo`,
      SPHttpClient.configurations.v1,
      { headers: { 'Accept':'application/json;odata=nometadata' } }
    );
    const data = await response.json();
    return data.FormDigestValue;
  }

  public async createJob(input: Partial<IJob>): Promise<number> {
    const list = 'GM Jobs';
    const map = await this.fieldMap(list);
    const payload: any = {};

    const set = (title:string, value:any) => {
      const internal = map[title];
      if (internal && value !== undefined) payload[internal] = value;
    };

    set('Title', input.title || `${input.customer || ''} - ${input.site || ''}`.trim());
    set('Customer', input.customer || '');
    set('Site', input.site || '');
    set('Scheduled Date', input.scheduledDate);
    set('Job Description', input.jobDescription || '');
    set('Crew', input.crew || '');
    set('Billing Type', input.billingType || '');
    set('Invoice Value', Number(input.invoiceValue || 0));
    set('Job Cost', Number(input.jobCost || 0));
    set('PO Number', input.poNumber || '');
    set('Status', input.status || 'Scheduled');
    set('PSI Number', input.psiNumber || '');

    const digest = await this.digest();
    const response = await this.context.spHttpClient.post(
      `${this.siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(list)}')/items`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          'Accept':'application/json;odata=nometadata',
          'Content-Type':'application/json;odata=nometadata',
          'X-RequestDigest': digest
        },
        body: JSON.stringify(payload)
      }
    );
    if (!response.ok) throw new Error(`Could not create job: ${response.statusText}`);
    const created = await response.json();
    return created.Id;
  }

  public async updateJobStatus(jobId: number, status: string): Promise<void> {
    const list = 'GM Jobs';
    const map = await this.fieldMap(list);
    const payload: any = {};
    payload[map['Status']] = status;

    const now = new Date().toISOString();
    if (status === 'Completed' && map['Completed Date']) payload[map['Completed Date']] = now;
    if (status === 'Invoiced' && map['Invoiced Date']) payload[map['Invoiced Date']] = now;

    const digest = await this.digest();
    const response = await this.context.spHttpClient.post(
      `${this.siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(list)}')/items(${jobId})`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          'Accept':'application/json;odata=nometadata',
          'Content-Type':'application/json;odata=nometadata',
          'IF-MATCH':'*',
          'X-HTTP-Method':'MERGE',
          'X-RequestDigest': digest
        },
        body: JSON.stringify(payload)
      }
    );
    if (!response.ok) throw new Error(`Could not update job: ${response.statusText}`);
  }
}

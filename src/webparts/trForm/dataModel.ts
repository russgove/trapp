import * as moment from 'moment';
export enum modes {
    NEW,
    EDIT,
    DISPLAY
}
export class peopleSearchResults {
    constructor(
        public PreferredName: string,
        public Department: string,
        public JobTitle: string,
        public PictureURL: string,
        public OfficeNumber: string
    ) { }
}
export class TR {

    public constructor() {
        this.Id = -1;
        this.ParentTR = null;
        this.ParentTRId = null;
        this.Title = '';
        this.CER = "";
        this.InitiationDate = moment(new Date()).toISOString();
        this.TRDueDate = moment(new Date()).toISOString();
        this.ActualStartDate = moment(new Date()).toISOString();
        this.ActualCompletionDate = moment(new Date()).toISOString();
        this.RequestorId = null;
        this.RequestorName = null;
        this.EstimatedHours = 0;
        this.Site = "";
        this.MailBox = "";
        this.TRPriority = "";
        this.Customer = "";
        this.Status = "";
        this.ApplicationTypeId = 0;
        this.EndUseId = 0;
        this.TitleArea = "";
        this.DescriptionArea = "";
        this.SummaryArea = "";
        this.WorkTypeId = 0;

    }
    public Id: number;
    public ParentTR: string;
    public ParentTRId: number;
    public Title: string;
    public CER: string;
    public InitiationDate: string;
    public TRDueDate: string;
    public ActualStartDate: string;
    public ActualCompletionDate: string;
    public RequestorId: number;
    public RequestorName:string;
    public EstimatedHours: number;
    public Site: string;
    public MailBox: string;
    public TRPriority: string;
    public Customer: string;
    public Status: string;
    public ApplicationTypeId: number;
    public EndUseId: number;
    public TitleArea: string;
    public DescriptionArea: string;
    public SummaryArea: string;
    public WorkTypeId: number;

}
export class WorkType {
    public constructor(
        public id: string,
        public workType: string) { }

}

export class ApplicationType {
    public constructor(
        public id: string,
        public applicationType: string,

        public workTypeIds: number[]

    ) { }

}
export class EndUse {
    public constructor(
        public id: string,
        public endUse: string,
        public applicationTypeId: number,
    ) { }

}
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
        this.Id = null;
        this.ParentTR = null;
        this.ParentTRId = null;
        this.Title = null;
        this.CER = null;
        this.InitiationDate =null; //moment(new Date()).toISOString();
        this.TRDueDate = null;//;moment(new Date()).toISOString();
        this.ActualStartDate = null;//moment(new Date()).toISOString();
        this.ActualCompletionDate = null;//= moment(new Date()).toISOString();
        this.RequestorId = null;
        this.RequestorName = null;
        this.EstimatedHours = null;
        this.Site = null;
 
        this.TRPriority =null;
        this.Customer = null;
        this.Status = null;
        this.ApplicationTypeId = null;
        this.EndUseId = null;
        this.TitleArea = null;
        this.DescriptionArea = null;
        this.SummaryArea = null;
                this.TestParamsArea = null;
        this.WorkTypeId = null;
           this.TechSpecId = [];

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
    public TRPriority: string;
    public Customer: string;
    public Status: string;
    public ApplicationTypeId: number;
    public EndUseId: number;
    public TitleArea: string;
    public DescriptionArea: string;
    public SummaryArea: string;
    public TestParamsArea: string;
    public WorkTypeId: number;
    public TechSpecId:Array<number>

}
export class WorkType {
    public constructor(
        public id: string,
        public workType: string) { }

}
export class User{
    public constructor(
        public id: number,
        public title: string,
        public position:string,
        public department:string,
        ) { }
    
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
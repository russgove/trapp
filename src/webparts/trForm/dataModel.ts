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
        this.RequestDate = null; //moment(new Date()).toISOString();
        this.RequiredDate = null;//;moment(new Date()).toISOString();
        this.ActualStartDate = null;//moment(new Date()).toISOString();
        this.ActualCompletionDate = null;//= moment(new Date()).toISOString();
        this.RequestorId = null;
        this.RequestorName = null;
        this.EstManHours = null;
        this.Site = null;

        this.TRPriority = null;
        this.CustomerId = null;
        this.TRStatus = null;
        this.ApplicationTypeId = null;
        this.EndUseId = null;
        this.RequestTitle = null;
        this.Description = null;
        this.Summary = null;
        this.TestingParameters = null;
        this.FormulaeArea = null;
        this.WorkTypeId = null;
        this.TechSpecId = null;

    }
    public Id: number;
    public ParentTR: string;
    public ParentTRId: number;
    public Title: string;
    public CER: string;
    public RequestDate: string;//InitiationDate
    public RequiredDate: string;//TRDueDate
    public ActualStartDate: string;
    public ActualCompletionDate: string;
    public RequestorId: number;
    public RequestorName: string;
    public EstManHours: number;//Edtimated hours
    public Site: string;
    public TRPriority: string;//TRPriority
    public CustomerId:number;
    public TRStatus: string;
    public ApplicationTypeId: number;
    public EndUseId: number;
    public RequestTitle: string;//TitleArea
    public Description: string;//DescriptionArea
    public Summary: string;//SummaryArea
    public FormulaeArea: string;
    public TestingParameters: string;//TestParamsArea
    public WorkTypeId: number;
    public TechSpecId: Array<number>;


}
export class WorkType {
    public constructor(
        public id: string,
        public workType: string) { }

}
export class Customer {
    public constructor(
        public id: number,
        public title: string,

    ) { }

}
export class User {
    public constructor(
        public id: number,
        public title: string,
        public position: string,
        public department: string,
    ) { }

}
export class ApplicationType {
    public constructor(
        public id: number,
        public applicationType: string,

        public workTypeIds: number[]

    ) { }

}
export class EndUse {
    public constructor(
        public id: number,
        public endUse: string,
        public applicationTypeId: number,
    ) { }

}
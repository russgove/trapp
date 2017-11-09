import { IPersonaProps } from 'office-ui-fabric-react/lib/Persona';
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
export class TRDocument {
    public constructor(
        public id:number,
        public title: string,
        public serverRalativeUrl: string,
        public length: number,
        public fileName: string,
        public majorVersion: number,
        public monorVersion: number,
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
        this.ActualManHours=null;
        this.Site = null;

        this.TRPriority = null;
        this.CustomerId = null;
        this.TRStatus = null;
        this.ApplicationTypeId = null;
        this.EndUseId = null;
        this.RequestTitle = null;
        this.Description = null;
        this.Summary = null;
        this.SummaryNew = null;
        this.TestingParameters = null;
        this.Formulae = null;
        this.WorkTypeId = null;
        this.TRAssignedToId = null;
        this.StaffCC = null;
        this.PigmentsId = null;
        this.TestsId = null;

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
    public ActualManHours: number;//Edtimated hours
    public Site: string;
    public TRPriority: string;//TRPriority
    public CustomerId: number;
    public TRStatus: string;
    public ApplicationTypeId: number;
    public EndUseId: number;
    public RequestTitle: string;//TitleArea
    public Description: string;//DescriptionArea
    // User request from carol-lynn:Is it possible for the “summary” part (free text area) to have a date automatically generated whenever someone updates it?
    //Just like how entering notes in SAP works. 
    // so the Summary will be the text we get from sharepoint , and when we save i will append SummaryNew to Summary
    public Summary: string;//SummaryArea-- not updateable in UI
    public SummaryNew: string;// Updatable in UI, will be appended to Summary before saving.
    public Formulae: string;
    public TestingParameters: string;//TestParamsArea
    public WorkTypeId: number;
    public TRAssignedToId: Array<number>;
    public StaffCC: Array<IPersonaProps>;
    public PigmentsId: Array<number>;
    public TestsId: Array<number>;



}
export class SetupItem { // just used in the display to show tests grouped by property
    public Title: string;
    public RichText: string;
    public PlainText: string;

}
export class TRFieldDefinition{
    public Title:string;
    public InternalName:string;
    public Description:string;
}
export class DisplayPropertyTest { // just used in the display to show tests grouped by property
    public property: string;
    public test: string;
    public testid: number;
}
export class Pigment {
    public manufacturer: string;
    public constructor(
        public id: number,
        public title: string,
        public isActive:string

    ) { this.manufacturer = null; }

}
export class WorkType {
    public constructor(
        public id: string,
        public workType: string) { }

}
export class Test {
    public constructor(
        public id: number,
        public title: string) { }

}
export class PropertyTest {
    public property: string;
    public constructor(
        public id: number,
        public applicationTypeid: number,
        public endUseIds: Array<number>,
        public testIds: Array<number>) {
        this.property = "";
    }

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
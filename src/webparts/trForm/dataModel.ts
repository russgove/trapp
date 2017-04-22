export class TR {
    public Id: number;
    public Title: string;
    public CER: string;
    public InitiationDate: string;
    public TRDueDate: string;
    public ActualStartDate: string;
    public ActualCompletionDate: string;
    public Requestor: string;
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
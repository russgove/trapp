export class TimeSpent {
    public TechnicalSpecialist: number;
    public TR: TechnicalRequest;
    public WeekEndingDate: Date;
    public HoursSpent: number;
    public Id:number;

}
export class TechnicalRequest{
    public id:number;
    public title:string;
    public status:string;
    public requiredDate:Date;
    
}
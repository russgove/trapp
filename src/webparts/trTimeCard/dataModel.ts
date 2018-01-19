
/**
 * Contains all information used to render the TimeSPen grid. This contains info from bot the TimeSPent table and the TR Table
 * 
 * @export
 * @class TimeSpent
 */
export class TimeSpent {
    public technicalSpecialist: number;
    public weekEndingDate: Date;
    public hoursSpent: number;// this is whats recored already in the hours spent
    public newHoursSpent: number; // this is whats added by the user
    public tsId: number;

    public trId: number;
    public trTitle: string;
    public trRequestTitle: string;
    public trStatus: string;
    public trPriority: string;
    public trRequiredDate: Date;

}

/**
 * Contains TR metadata to be added to a TimeSPent record
 * 
 * @export
 * @class TechnicalRequest
 */
export class TechnicalRequest {
    public trId: number;
    public title: string;
    public requestTitle: string;
    public status: string;
    public requiredDate: Date;
    public priority: string;


}
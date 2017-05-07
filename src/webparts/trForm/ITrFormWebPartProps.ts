import { modes } from "./dataModel";
export interface ITrFormWebPartProps {
  mode: modes;
  technicalRequestListName: string,
  applicationTYpeListName: string,
  endUseListName: string,
  workTypeListName: string,
  custonersListName: string,
  pigmentListName: string,
  propertyTestListName: string, // do i need this, can i just get with the PropertyTest via expand
  partyListName:string, //Customers
  propertyListName: string, 
  testListName: string, 
}

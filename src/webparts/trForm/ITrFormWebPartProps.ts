import { modes } from "./dataModel";
export interface ITrFormWebPartProps {
  mode: modes;
  technicalRequestListName: string;
  applicationTYpeListName: string;
  endUseListName: string;
  workTypeListName: string;
  nextNumbersListName:string;
 setupListName:string;
 
  custonersListName: string;
  pigmentListName: string;
  propertyTestListName: string; // do i need this, can i just get with the PropertyTest via expand
  partyListName:string; //Customers
 // propertyListName: string,  
  testListName: string;
  trDocumentsListName:string;
  searchPath:string; // path passed to the search engine when searchng for trs
  defaultSite:string;// value to put into the site column on new trs
  enableEmail:boolean;// to disable sending emails while testing
    editFormUrlFormat:string;//
  displayFormUrlFormat:string;
}

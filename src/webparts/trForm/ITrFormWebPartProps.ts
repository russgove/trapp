import { modes } from "./dataModel";
export interface ITrFormWebPartProps {
  mode: modes;
  technicalRequestListName: string;
  applicationTYpeListName: string;
  endUseListName: string;
  workTypeListName: string;
  nextNumbersListName: string;
  setupListName: string;

  custonersListName: string;
  pigmentListName: string;
  propertyTestListName: string; // do i need this, can i just get with the PropertyTest via expand
  partyListName: string; //Customers
  // propertyListName: string,  
  testListName: string;
  trDocumentsListName: string;
  searchPath: string; // path passed to the search engine when searchng for trs
  defaultSite: string;// value to put into the site column on new trs
  enableEmail: boolean;// to disable sending emails while testing
  editFormUrlFormat: string;//
  displayFormUrlFormat: string;
  delayPriorToSettingCKEditor:number;
  emailSuffix:string; // when searching for People in staffCC we only rturn users with emails ending in emailsuffix
  visitorsGoupdName:string;// When we add a staffCC user gets added to this group so he can visit site
  ckeditorUrl:string; //path to load ckeditor from  (//cdn.ckeditor.com/4.6.2/full/ckeditor.js  OR our  cdn)
  documentIframeHeight:number; // heighht of the iframe that shows the document in the Documents Tab
  documentIframeWidth:number; // width of the iframe that shows the document in the Documents Tab
  workflowToTerminateOnChange:string; // the name of the workflow to terminate when an item changes ("Send TR Norifications")
}

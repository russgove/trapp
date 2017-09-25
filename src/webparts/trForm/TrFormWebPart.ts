
import { IPersonaProps } from 'office-ui-fabric-react/lib/Persona';
import * as React from 'react';
import * as ReactDom from 'react-dom';
import pnp from "sp-pnp-js";
import * as moment from "moment";
import { SearchQuery, SearchResults, SortDirection, EmailProperties, Items } from "sp-pnp-js";
import { Version, UrlQueryParameterCollection } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration, PropertyPaneDropdown, PropertyPaneTextField, PropertyPaneCheckbox
} from '@microsoft/sp-webpart-base';
import {
  Environment,
  EnvironmentType
} from '@microsoft/sp-core-library';
import { TRDocument, SetupItem, Test, PropertyTest, Pigment, TR, WorkType, ApplicationType, EndUse, modes, User, Customer } from "./dataModel";
import * as strings from 'trFormStrings';
import * as _ from 'lodash';
import TrForm from './components/TrForm';
import { ITrFormProps } from './components/ITrFormProps';
import { ITRFormState } from './components/ITRFormState';
import { ITrFormWebPartProps } from './ITrFormWebPartProps';


/**
 * Webpart used to display the new and edit forms for Technical Requests.
 * 
 * @export
 * @class TrFormWebPart
 * @extends {BaseClientSideWebPart<ITrFormWebPartProps>}
 */
export default class TrFormWebPart extends BaseClientSideWebPart<ITrFormWebPartProps> {
  private tr: TR;
  private childTRs: Array<TR>;
  private reactElement: React.ReactElement<ITrFormProps>;
  private trContentTypeID: string;

  public onInit(): Promise<void> {
    return super.onInit().then(_ => {
      pnp.setup({
        spfxContext: this.context,
      });
      //  this.loadData();
    });
  }

  /**
   * Utility method to move all the data from a listitem we got from the TR list to a TR record
   * 
   * @private
   * @param {TR} tr The TR to add the fields to.
   * @param {*} item the SPLIst item from teh Technical requests list
   * 
   * @memberof TrFormWebPart
   */
  private moveFieldsToTR(tr: TR, item: any) {
    tr.Id = item.Id;
    tr.ActualCompletionDate = item.ActualCompletionDate;
    tr.ApplicationTypeId = item.ApplicationTypeId;
    tr.ActualStartDate = item.ActualStartDate;
    tr.CER = item.CER;
    tr.CustomerId = item.CustomerId;

    tr.RequiredDate = item.RequiredDate;

    tr.EstManHours = item.EstManHours;
    tr.RequestDate = item.RequestDate;
    tr.TRPriority = item.TRPriority;
    tr.RequestorId = item.RequestorId;
    if (item.Requestor) {
      tr.RequestorName = item.Requestor.Title;
    }
    tr.Site = item.Site;
    tr.TRStatus = item.TRStatus;
    tr.EndUseId = item.EndUseId;
    tr.WorkTypeId = item.WorkTypeId;
    tr.Title = item.Title;
    tr.RequestTitle = item.RequestTitle;
    tr.Formulae = item.Formulae;
    tr.Description = item.Description;
    tr.Summary = item.Summary;
    tr.TestingParameters = item.TestingParameters;
    tr.ParentTRId = item.ParentTRId;
    if (item.ParentTR) {
      tr.ParentTR = item.ParentTR.Title;
    }
    tr.TRAssignedToId = item.TRAssignedToId;
    tr.StaffCC = this.getStaffCCFromTR(item);
    tr.PigmentsId = item.PigmentsId;
    tr.TestsId = item.TestsId;
  }
  /**
 * Method to extract Personas from the STAfcc fields on a TR
 * 
 * @param {item} a tr getch throuh the rest api expanding the staffcc fields
 * @returns {Promise<TR>}  A Promise for the TR record
 * 
 * @memberof TrFormWebPart
 */
  public getStaffCCFromTR(item: any): Array<IPersonaProps> {

    let personas: Array<IPersonaProps> = [];

    if (item.StaffCC) {
      for (let staffcc of item.StaffCC) {

        personas.push({
          primaryText: staffcc["Title"],
          secondaryText: staffcc["JobTitle"],
          tertiaryText: staffcc["Department"],
          optionalText: staffcc["EMail"],
          //imageUrl:result["PictureURL"], cannot expand Picure when I join TR to site users list, would need to doubleback and get thes
          id: staffcc['Id']
        });
      }
    }
    return personas;
  }
  /**
   * Method to fetch a TR from the Technical Request list
   * 
   * @param {number} id The id of the TR to fetch
   * @returns {Promise<TR>}  A Promise for the TR record
   * 
   * @memberof TrFormWebPart
   */
  public fetchTR(id: number): Promise<TR> {
    let fields = "*,ParentTR/Title,Requestor/Title";
    return pnp.sp.web.lists.getByTitle(this.properties.technicalRequestListName).items.getById(id).expand("ParentTR,Requestor").select(fields).get()
      .then((item) => {
        let tr = new TR();
        this.moveFieldsToTR(tr, item);
        return tr;
      });
  }

  /**
   * A method to fetch the WopiFrameURL for a Document in the TR Documents library.
   * This url is used to display the document in the iframs
   * @param {number} id the listitem id of the document in the TR Document Libtry
   * @param {number} mode  The displayMode in the retuned url (display, edit, etc.)
   * @returns {Promise<string>} The url used to display the document in the iframe
   * 
   * @memberof TrFormWebPart
   */
  public fetchDocumentWopiFrameURL(id: number, mode: number): Promise<string> {
    let fields = "*,ParentTR/Title,Requestor/Title";
    return pnp.sp.web.lists.getByTitle(this.properties.trDocumentsListName).items.getById(id).getWopiFrameUrl(mode).then((item) => {
      return item;
    });
  }

  /**
   * Method to fetch All child TRS for the selected TR
   * 
   * @param {number} id The ID of the TR to fetch children for
   * @returns {Promise<Array<TR>>} An array of TRs that are childrent of the selected TRS. This is just a self-referncing lookup column.
   * 
   * @memberof TrFormWebPart
   */
  public fetchChildTR(id: number): Promise<Array<TR>> {
    let fields = "*,ParentTR/Title,Requestor/Title";
    return pnp.sp.web.lists.getByTitle(this.properties.technicalRequestListName).items.filter("ParentTR eq " + id).expand("ParentTR,Requestor").select(fields).get()
      .then((items) => {
        let childTrs = new Array<TR>();

        for (const item of items) {
          let childtr: TR = new TR();
          this.moveFieldsToTR(childtr, item);
          childTrs.push(childtr);
        }
        return childTrs;
      });

  }


  /**
   *  An accesser indicating whether or not the current page is in design mode.
   * Used to turn the ribbon off when not in edit mode
   * 
   * @returns {boolean} 
   * 
   * @memberof TrFormWebPart
   */
  public inDesignMode(): boolean {
    if (document.getElementById("MSOLayout_InDesignMode")) {
      console.log(
        document.getElementById("MSOLayout_InDesignMode")
      );
      if (document.getElementById("MSOLayout_InDesignMode").innerText === "1") {
        return true;
      }
      else {
        console.log('document.getElementById("MSOLayout_InDesignMode") is null');
        return false;
      }
    }
    else {
      console.log("document.getElementById(MSOLayout_InDesignMode) is false");
      return false;
    }

  }

  /**
   * Gets all the documents associated with a selected TR
   * 
   * @param {number} id The TR Id to get documents for
   * @param {*} [batch] The odata batch to execute the call in (from pnp.SP.createBatch). If not present the request wont be batched.
   * @returns {Promise<Array<TRDocument>>} The TRDocuments for the selectd Tr
   * 
   * @memberof TrFormWebPart
   */
  public getDocuments(id: number, batch?: any): Promise<Array<TRDocument>> {
    let docfields = "Id,Title,File/ServerRelativeUrl,File/Length,File/Name,File/MajorVersion,File/MinorVersion";
    let docexpands = "File";

    let command: Items = pnp.sp.web.lists
      .getByTitle(this.properties.trDocumentsListName)
      .items.filter("TR eq " + id)
      .expand(docexpands)
      .select(docfields);
    if (batch) {
      command.inBatch(batch);
    }
    return command.get().then((items) => {
      let docs: Array<TRDocument> = [];

      for (const item of items) {
        let trDoc: TRDocument = new TRDocument(item.Id, item.Title, item.File.ServerRelativeUrl, item.File.Length, item.File.Name, item.File.MajorVersion, item.File.MinorVersion);
        docs.push(trDoc);
      }
      return docs;
    });

  }

  /**
   * Renders the react Form.
   * Fetches the initial data and renders the react compmonent, then fetches all the ancilliary dlookup data 
   * and calls forceUpdate on the component to push down the additional lookup data
   * 
   * @memberof TrFormWebPart
   */
  public render(): void {
    // hide the ribbon
    //if (!this.inDesignMode())

    if (document.getElementById("s4-ribbonrow")) {
      document.getElementById("s4-ribbonrow").style.display = "none";
    }

    let formProps: ITrFormProps = {
      save: this.save.bind(this),
      fetchChildTr: this.fetchChildTR.bind(this),
      fetchTR: this.fetchTR.bind(this),
      fetchDocumentWopiFrameURL: this.fetchDocumentWopiFrameURL.bind(this),
      cancel: this.cancel.bind(this),
      TRsearch: this.TRsearch.bind(this),
      uploadFile: this.uploadFile.bind(this),
      getDocuments: this.getDocuments.bind(this),
      peopleSearch: this.PeopleSearch.bind(this),
      ensureUsersInPersonas: this.ensureUsersInPersonas.bind(this),

      customers: [],
      initialState: null,
      techSpecs: [],
      requestors: [],
      mode: this.properties.mode,
      workTypes: [],
      applicationTypes: [],
      endUses: [],
      pigments: [],
      tests: [],
      propertyTests: [],
      ckeditorUrl: this.properties.ckeditorUrl,
      delayPriorToSettingCKEditor: this.properties.delayPriorToSettingCKEditor,
      ckeditorConfig:{}


    };
    let formState: ITRFormState = {
      tr: new TR(),
      childTRs: [],
      errorMessages: [],
      isDirty: false,
      showTRSearch: false,
      documentCalloutVisible: false,
      documents: [],
      documentCalloutTarget: null,
      documentCalloutIframeUrl: null
    };
    let batch = pnp.sp.createBatch();

    pnp.sp.web.lists.getByTitle(this.properties.setupListName).items.filter("Title eq 'ckeditorConfig'").inBatch(batch).getAs<SetupItem[]>()
    .then((setupItems) => {
      debugger;
      formProps.ckeditorConfig = JSON.parse(setupItems[0].PlainText)
    })
    .catch((error) => {
      console.log("ERROR, An error occured fetching and parsing ckeditorConfig " + this.properties.setupListName);
      console.log(error.message);

    });
    pnp.sp.web.lists.getByTitle(this.properties.endUseListName).items.inBatch(batch).get()
      .then((items) => {
        formProps.endUses = _.map(items, (item) => {
          return new EndUse(item["Id"], item["Title"], item["ApplicationTypeId"]);
        });
      })
      .catch((error) => {
        console.log("ERROR, An error occured fetching 'End uses' from list named " + this.properties.endUseListName);
        console.log(error.message);

      });

    pnp.sp.web.lists.getByTitle(this.properties.applicationTYpeListName).items.inBatch(batch).get()
      .then((items) => {

        formProps.applicationTypes = _.map(items, (item) => {
          return new ApplicationType(item["Id"], item["Title"], item["WorkTypesId"]);
        });

      })
      .catch((error) => {
        console.log("ERROR, An error occured fetching 'Application Types' from list named " + this.properties.applicationTYpeListName);
        console.log(error.message);

      });
    var queryParameters = new UrlQueryParameterCollection(window.location.href);

    if (this.properties.mode !== modes.NEW) {
      if (queryParameters.getValue("Id")) {
        const id: number = parseInt(queryParameters.getValue("Id"));
        let fields = "*,WorkType/Title,ParentTR/Title,Requestor/Title,Customer/Title,TRAssignedTo/Title,TRAssignedTo/Id,StaffCC/EMail,StaffCC/Title,StaffCC/Name,StaffCC/JobTitle,StaffCC/Department,StaffCC/Id";
        let expands = "ParentTR,Requestor,Customer,TRAssignedTo,WorkType,StaffCC";
        // get the requested tr
        pnp.sp.web.lists.getByTitle(this.properties.technicalRequestListName).items.getById(id).expand(expands).select(fields).inBatch(batch).get()

          .then((item) => {
            formState.tr = new TR();
            this.moveFieldsToTR(formState.tr, item);
            if (item.Customer) {// single value lookup
              formProps.customers.push(new Customer(item.CustomerId, item.Customer.Title));
            }
            if (item.TRAssignedTo) {// multi value lookup
              for (let assignee of item.TRAssignedTo) {
                formProps.techSpecs.push(new User(assignee["Id"], assignee["Title"]));
              }
            }
            if (item.Requestor) {// single value lookup
              formProps.requestors.push(new User(item.RequestorId, item.Requestor.Title));
            }
            if (item.WorkType) {// single value lookup
              formProps.workTypes.push(new WorkType(item.WorkTypeId, item.WorkType.Title));
            }
          })
          .catch((error) => {
            console.log("ERROR, An error occured fetching the listitem  from list named " + this.properties.technicalRequestListName);
            console.log(error.message);

          });
        // get the Child trs
        pnp.sp.web.lists.getByTitle(this.properties.technicalRequestListName).items.filter("ParentTRId eq " + id).expand(expands).select(fields).inBatch(batch).get()

          .then((items) => {
            // this may resilve befor we get the mainn tr, so jyst stash  away for now.
            for (const item of items) {
              let childtr: TR = new TR();
              this.moveFieldsToTR(childtr, item);
              formState.childTRs.push(childtr);
            }
          })
          .catch((error) => {
            console.log("ERROR, An error occured fetching child trs  from list named " + this.properties.technicalRequestListName);
            console.log(error.message);

          });
        // get the Documents

        this.getDocuments(id, batch).then((docs) => {
          formState.documents = docs;
        }).catch((error) => {
          console.log("ERROR, An error occured fetching Documents  from list named " + this.properties.trDocumentsListName);
          console.log(error.message);
        });
      }
      else {
        console.log("ERROR, Id not specified with Display or Edit form");
      }
    }
    else {
      formState.tr.Site = this.properties.defaultSite;
      pnp.sp.web.currentUser.inBatch(batch).get().then((user) => {

        formState.tr.RequestorId = user["Id"];
        formState.tr.RequestorName = user["Title"];
      })
        .catch((err) => {
          console.log("unable to fetch current user");
        });
      pnp.sp.web.lists.getByTitle(this.properties.nextNumbersListName).items.select("Id,NextNumber").filter("CounterName eq 'RequestId'").orderBy("Title").top(5000).inBatch(batch).get()// get the lookup info
        .then((items) => {
          if (items.length != 1) {
            console.log("multiple next numbers found");
          }
          else {
            let nextNumber: number = items[0]["NextNumber"];
            nextNumber++;
            formState.tr.Title = this.properties.defaultSite + nextNumber;

            pnp.sp.web.lists.getByTitle(this.properties.nextNumbersListName).items.getById(items[0].Id)
              .update({ "NextNumber": nextNumber }).then((results) => {
                console.log("next number not increment to " + nextNumber);
              }).catch((err) => {
                alert("next number not incremented-- please try again");
              });
          }
        }).catch((err) => {

          console.log("next number not increment to");
        });
    }


    batch.execute().then((value) => {// execute the batch to get the item being edited and info REQUIRED for initial display
      formProps.initialState = formState;
      this.reactElement = React.createElement(TrForm, formProps);
      var formComponent: TrForm = ReactDom.render(this.reactElement, this.domElement) as TrForm;//render the component
      window.onbeforeunload = function (e) {
        debugger;

        if (formComponent.state.isDirty) {
          var dialogText = "You have unsaved changes, are you sure you want to leave?";
          e.returnValue = dialogText;
          return dialogText;

        }
      };
      if (Environment.type === EnvironmentType.ClassicSharePoint) {
        const buttons: NodeListOf<HTMLButtonElement> = this.domElement.getElementsByTagName('button');
        if (buttons && buttons.length) {
          for (let i: number = 0; i < buttons.length; i++) {
            if (buttons[i]) {
              /* tslint:disable */
              // Disable the button onclick postback
              buttons[i].onclick = function () { return false; };
              /* tslint:enable */
            }
          }
        }
      }
      let batch2 = pnp.sp.createBatch(); // create a second batch to get the lookup columns
      const requestorsGroupName = "TR " + this.context.pageContext.web.title + " Requestors";
      pnp.sp.web.siteGroups.getByName(requestorsGroupName).users.orderBy("Title").inBatch(batch2).get()
        .then((items) => {
          let requestors: Array<User> = _.map(items, (item) => {
            return new User(item["Id"], item["Title"]);
          });
          formProps.requestors = _.unionWith(requestors, formProps.requestors, (a, b) => { return a.id === b.id; });//_.union

        })
        .catch((error) => {
          console.log("ERROR, An error occured fetching Requestors from group " + requestorsGroupName);
          console.log(error.message);

        });

      const techspecGroupName = "TR " + this.context.pageContext.web.title + " Tech Specialists";
      pnp.sp.web.siteGroups.getByName(techspecGroupName).users.orderBy("Title").inBatch(batch2).get()



        .then((items) => {
          let techSpecs: Array<User> = _.map(items, (item) => {
            return new User(item["Id"], item["Title"]);
          });
          formProps.techSpecs = _.unionWith(techSpecs, formProps.techSpecs, (a, b) => { return a.id === b.id; });//_.union

        })
        .catch((error) => {
          console.log("ERROR, An error occured fetching Tech Specialists from group " + techspecGroupName);
          console.log(error.message);

        });
      let customerFields = "Id,Title";
      pnp.sp.web.lists.getByTitle(this.properties.partyListName).items.select(customerFields).filter("IsActive eq 'Yes'").orderBy("Title").top(5000).inBatch(batch2).get()// get the lookup info
        .then((items) => {
          let customers: Array<Customer> = _.map(items, (item) => {
            return new Customer(item["Id"], item["Title"]);
          });
          // add the one from the tr if not present
          if (formProps.customers.length > 0 &&
            _.find(customers, (c) => { return c.id === formProps.customers[0].id; }) == null) {

            customers.push(formProps.customers[0]);
          }
          formProps.customers = customers;
        })
        .catch((error) => {
          console.log("ERROR, An error occured fetching 'Customers' from list " + this.properties.partyListName);
          console.log(error.message);
        });

      let workTypesFields = "Id,Title";
      pnp.sp.web.lists.getByTitle(this.properties.workTypeListName).items.filter("IsActive eq 'Yes'").inBatch(batch2).get()
        .then((items) => {
          let workTypes: Array<WorkType> = _.map(items, (item) => {
            return new WorkType(item["Id"], item["Title"]);
          });
          // add the one from the tr if not present
          if (formProps.workTypes.length > 0 &&
            _.find(workTypes, (wt) => { return wt.id === formProps.workTypes[0].id; }) == null) {
            workTypes.push(formProps.workTypes[0]);
          }
          formProps.workTypes = workTypes;
        })
        .catch((error) => {
          console.log("ERROR, An error occured fetching 'Work Types' from list named " + this.properties.workTypeListName);
          console.log(error.message);

        });
      let pigmentFields = "Id,Title,IsActive,Manufacturer/Title";
      let pigmentExpands = "Manufacturer";
      pnp.sp.web.lists.getByTitle(this.properties.pigmentListName).items.select(pigmentFields).expand(pigmentExpands).top(5000).inBatch(batch2).get()// get the lookup info
        .then((items) => {
          formProps.pigments = _.map(items, (item) => {
            let p: Pigment = new Pigment(item["Id"], item["Title"], item["IsActive"]);
            if (item["Manufacturer"]) {
              p.manufacturer = item["Manufacturer"]["Title"];
            }
            return p;
          });

        })
        .catch((error) => {
          console.log("ERROR, An error occured fetching 'Pigments' from list " + this.properties.pigmentListName);
          console.log(error.message);
        });
      let testFields = "Id,Title";
      pnp.sp.web.lists.getByTitle(this.properties.testListName).items.select(testFields).top(5000).inBatch(batch2).get()// get the lookup info
        .then((items) => {
          formProps.tests = _.map(items, (item) => {
            let t: Test = new Test(item["Id"], item["Title"]);
            return t;
          });
        })
        .catch((error) => {
          console.log("ERROR, An error occured fetching 'Pigments' from list " + this.properties.pigmentListName);
          console.log(error.message);
        });
      let propertyTestFields = "*,Property/Title";
      let propertyTestExpands = "Property";
      pnp.sp.web.lists.getByTitle(this.properties.propertyTestListName).items.select(propertyTestFields).expand(propertyTestExpands).top(5000).inBatch(batch2).get()// get the lookup info
        .then((items) => {
          formProps.propertyTests = _.map(items, (item) => {
            let pt: PropertyTest = new PropertyTest(item["Id"] as number, item["ApplicationTypeId"] as number, item["EndUseId"] as Array<number>, item["TestId"] as Array<number>);
            if (item["Property"]) {
              pt.property = item["Property"]["Title"];
            }
            return pt;
          });
        })
        .catch((error) => {
          console.log("ERROR, An error occured fetching 'PropertyText' from list " + this.properties.propertyTestListName);
          console.log(error.message);
        });
      batch2.execute().then(() => {
        //  formComponent.props = formProps; this did not work
        formComponent.props.customers = formProps.customers;
        formComponent.props.pigments = formProps.pigments;
        formComponent.props.tests = formProps.tests;
        formComponent.props.propertyTests = formProps.propertyTests;
        formComponent.props.techSpecs = formProps.techSpecs;
        formComponent.props.requestors = formProps.requestors;
        formComponent.props.workTypes = formProps.workTypes;
        formComponent.forceUpdate();
      });
    }
    );

  }
  protected delay(ms: number) {
    return new Promise(resolve => setTimeout(resolve, ms));
  }
  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneDropdown('mode', {
                  label: strings.ModeFieldLabel,
                  options: [
                    { text: "New", key: modes.NEW },
                    { text: "Edit", key: modes.EDIT },
                    { text: "Display", key: modes.DISPLAY },
                  ]
                }),
                PropertyPaneTextField('defaultSite', {
                  label: "Default Site",
                  description: " The value to place in the 'site' field on new TR's"
                }),
                PropertyPaneTextField('searchPath', {
                  label: "Search Path",
                  description: "The path passed to the search engine when searching for TR's"
                }),
                PropertyPaneTextField('editFormUrlFormat', {
                  label: "Edit Form Url format",
                  description: "USed to format the link to the edit form sent in emails"
                }),
                PropertyPaneTextField('displayFormUrlFormat', {
                  label: "Display Form Url format",
                  description: "USed to format the link to the display form sent in emails"
                }),
                PropertyPaneTextField('emailSuffix', {
                  label: "E-Mail Suffix",
                  description: "When searching for StaffCC only return people with email addresses ending with this"
                }),
                PropertyPaneTextField('visitorsGoupdName', {
                  label: "Visitors Group Name",
                  description: "When we add a StaffCC, the user gets added to this group so he can visit site"
                }),
                PropertyPaneCheckbox('enableEmail', {
                  text: "Enable sending emails to assignees and staff cc",
                }),
                PropertyPaneTextField('ckeditorUrl', {
                  label: "Url used to load CKEditor",
                  description: "CKEditor is the Roch text editor used oin the forms. It can be loaded from the public url(//cdn.ckeditor.com/4.6.2/full/ckeditor.js) or our cdn"
                }),


              ]
            },
            {
              groupName: "List Names",
              groupFields: [

                PropertyPaneTextField('technicalRequestListName', {
                  label: "Technical Requests list name",
                }),
                PropertyPaneTextField('applicationTYpeListName', {
                  label: "Application Types list name",
                }),
                PropertyPaneTextField('endUseListName', {
                  label: "End Uses list name",
                }),
                PropertyPaneTextField('workTypeListName', {
                  label: "Work Types list name",
                }),
                PropertyPaneTextField('setupListName', {
                  label: "Setup list name",
                }),
                PropertyPaneTextField('pigmentListName', {
                  label: "Pigments list name",
                }),
                PropertyPaneTextField('nextNumbersListName', {
                  label: "Next Numbers list name",
                }),
                PropertyPaneTextField('propertyListName', {
                  label: "Properties list name",
                }),
                PropertyPaneTextField('testListName', {
                  label: "Tests list name",
                }),
                PropertyPaneTextField('propertyTestListName', {
                  label: "Test Propeterties list name",
                }),
                PropertyPaneTextField('partyListName', {
                  label: "Customers list name",
                }),
                PropertyPaneTextField('trDocumentsListName', {
                  label: "TR Documents library name",
                }),
              ]
            }
          ]
        }
      ]
    };
  }
  public addUserToVisitorsGroup(userName: string) {
    pnp.sp.web.siteGroups.getByName(this.properties.visitorsGoupdName).users.add(userName).then((d) => {
    }).catch((err) => {
      console.log("error adding user to visitors group");
      console.log(err);

    });
  }
  public ensureUsersInPersonas(items: Array<IPersonaProps>): void {
    for (const item of items) {
      if (item.id === null) {
        pnp.sp.web.ensureUser(item.optionalText).then((result) => {
          item.id = result.data.Id.toString();
          this.addUserToVisitorsGroup(item.optionalText);

        }).catch((error) => {
          console.log("Error: failed to ensure user with email addrss " + item.optionalText);
        });
      }
    }

  }
  public PeopleSearch(filter: string, selectedItems?: Array<IPersonaProps>): Promise<Array<IPersonaProps>> {

    const query: SearchQuery = {

      Querytext: 'PreferredName:' + filter + '*',
      SourceId: 'b09a7990-05ea-4af9-81ef-edfab16c4e31', //people
      RowLimit: 15,
      SelectProperties: [
        'JobTitle',
        'WorkEmail',
        'PreferredName',
        'Department',
        'PictureURL',
        'Name',
        'AccountName'

      ],
      // SortList:[
      //   {Property:'PreferredName',Direction:SortDirection.Ascending}
      // ]


    };
    return pnp.sp.search(query).then((results: SearchResults) => {

      let personas: Array<IPersonaProps> = [];
      const suffix: string = this.properties.emailSuffix.toUpperCase();
      for (const result of results.PrimarySearchResults) {
        const email: string = result["WorkEmail"];

        if (_.findIndex(selectedItems, (si) => { return si.optionalText === result["WorkEmail"]; }) === -1 &&
          email != null &&
          email.toUpperCase().substr(-suffix.length) === suffix // endsWith
        ) {
          personas.push({
            primaryText: result["PreferredName"],
            secondaryText: result["JobTitle"],
            tertiaryText: result["Department"],
            optionalText: result["AccountName"],
            imageUrl: result["PictureURL"],
            id: null, // this needs to be set to the ID of the user in the sharepoint site. If user is selected we need to ensure user then add the ID
          });
        }
      }
      return personas;
    }).catch((e) => {
      debugger;
      console.log("peoplesearch thew error " + e);
      return null;
    });

  }
  /**
   * Calls sharepoint search to find TRS to be set as the parent TR.
   * The search path is set to only find items in the Technical Request List.
   * 
   * 
   * @param {string} searchText 
   * @returns {Promise<TR[]>} An array of TRS that match the seach text. Tese TRS do not have all metadata, just the data
   * we need to display the search results
   * 
   * @memberof TrFormWebPart
   */

  public TRsearch(searchText: string): Promise<TR[]> {

    //let queryText = "{0} Path:{1}* ContentTypeId:{2}*";
    let queryText = "{0} Path:{1}*";
    queryText = queryText
      .replace("{0}", searchText)
      .replace("{1}", this.properties.searchPath.split('{0}').join(this.context.pageContext.web.absoluteUrl));
    //.replace("{2}", this.trContentTypeID);
    let sq: SearchQuery = {
      Querytext: queryText,
      RowLimit: 50,
      SelectProperties: ["Title", "ListItemID", "RefinableString13", "RefinableString08", "RefinableString14", "TR-RequestTitle", "Description"],
      ///SortList: [{ Property: "PreferredName", Direction: SortDirection.Ascending }] arghhh-- not sortable
      Refiners: "RefinableString02,RefinableString03"
    };
    // refiners are in primarry query results reinemnet refiners
    console.log(sq);

    return pnp.sp.search(sq).then((results: SearchResults) => {
      let returnValue: Array<TR> = [];
      for (let sr of results.PrimarySearchResults) {
        const temp = sr as any;
        let tr: TR = new TR();
        tr.Id = temp.ListItemID;
        tr.Title = temp.Title;
        tr.CustomerId = temp.RefinableString08;
        tr.Site = temp.RefinableString14;
        tr.CER = temp.RefinableString13;
        tr.RequestTitle = temp["TR-RequestTitle"];
        tr.Description = temp["Description"];
        returnValue.push(tr);
      }
      return _.sortBy(returnValue, "Title");
    });
  }

  /**
   * Navigates to whatever URL was specified in the @Source=Querystring parameter.
   * 
   * @private
   * 
   * @memberof TrFormWebPart
   */
  private navigateToSource() {
    let queryParameters = new UrlQueryParameterCollection(window.location.href);
    let encodedSource = queryParameters.getValue("Source");
    if (encodedSource) {
      let source = decodeURIComponent(encodedSource);
      console.log("Navigating to source source uis " + source);
      window.location.href = source;
    }
    else {
      console.log("no  source staying on page");

    }
  }

  /**
   * Saves an updated TR back to Sharepoint , or adds a new one if no TR id is present,.
   * 
   * @private
   * @param {TR} tr A tr record with the data to be saved.
   * @param {Array<number>} orginalAssignees The list of people assigned to this TR before we started editting.
   * If there are assignees on the TR we are saving that are not in the orginalAssignees, we sent them and email
   * saying they have been added.
   * @param {string} originalStatus The Status of the TR before we started editing. If the New Status is completed and 
   * the old status was not completed , we email everyone in the StafCC list that the TR is now completed
   * @returns {Promise<any>} 
   * 
   * @memberof TrFormWebPart
   */
  private save(tr: TR, orginalAssignees: Array<number>, originalStatus: string): Promise<any> {
    // remove lookups
    let copy = _.clone(tr) as any;
    delete copy.RequestorName;
    delete copy.ParentTR;

    // reformat multivalue lookups for save
    let technicalSpecialists = (copy.TRAssignedToId) ? copy.TRAssignedToId : [];
    delete copy.TechSpecId;
    copy["TRAssignedToId"] = {};
    copy["TRAssignedToId"]["results"] = technicalSpecialists;
    console.log("reformatetd techSpecs for save");

    // staffcc is an array of IPersonaProps where the id field is the users ID in the user infomation list.
    // We need to convert this to StaffCCId/resulsts/ids to post back
    copy["StaffCCId"] = {};
    copy["StaffCCId"]["results"] = _.map(copy.StaffCC, (cc: IPersonaProps) => {

      return parseInt(cc.id);
    });
    delete copy.StaffCC;
    console.log("reformatetd staffcc for save");

    let TestsId = (copy.TestsId) ? copy.TestsId : [];
    delete copy.TestsId;
    copy["TestsId"] = {};
    copy["TestsId"]["results"] = TestsId;
    console.log("reformatetd tests for save");

    let PigmentsId = (copy.PigmentsId) ? copy.PigmentsId : [];
    delete copy.PigmentsId;
    copy["PigmentsId"] = {};
    copy["PigmentsId"]["results"] = PigmentsId;

    console.log("reformatetd pigments for save");
    // append the date and SummaryNew to Summary prior to save.
    if (copy.SummaryNew) {

      let today = moment(new Date()).format("DD-MMM-YYYY");
      if (copy.Summary) {
        copy.Summary = copy.Summary + "<br /><b>" + today + "</b><br />" + copy.SummaryNew;
      }
      else {
        copy.Summary = "<b>" + today + "</b><br />" + copy.SummaryNew;

      }

    }
    if (copy.hasOwnProperty("SummaryNew")) {
      delete copy.SummaryNew;
    }
    if (copy.Id !== null) {
      console.log("id is mot null will update");
      return pnp.sp.web.lists.getByTitle(this.properties.technicalRequestListName).items.getById(tr.Id).update(copy).then((item) => {
        console.log("Item sucessfully added, emailing asignnes");
        let newAssigneesPromise = this.emailNewAssignees(tr, orginalAssignees);
        console.log("emailling staff cc");
        var staffccPromise = this.emailStaffCC(tr, originalStatus);
        console.log("awaiting promises from emails");
        return Promise.all([newAssigneesPromise, staffccPromise])
          .then((a) => {

            console.log("emails sent continuing");
            let x = newAssigneesPromise;
            let y = staffccPromise;
            this.navigateToSource();// should stop here when on a form page  
            return tr;
          })
          .catch((err) => {

            console.log("error sending emails " + err);
          });
      });
    }
    else {
      console.log("id is  null will add");
      delete copy.Id;
      return pnp.sp.web.lists.getByTitle(this.properties.technicalRequestListName).items.add(copy).then((item) => {
        let newTR = new TR();

        newTR.Id = item.data.Id; // will be passed back toi component and component will set this to th eID NOT REALLY NEEDED
        newTR.TRAssignedToId = copy.TRAssignedToId.results;//used to email new assignees
        newTR.Title = copy.Title;
        // just makes debugging easier
        var newAssigneesPromise = this.emailNewAssignees(newTR, orginalAssignees);
        // var staffccPromise = this.emailStaffCC(newTR, originalStatus);
        return newAssigneesPromise.then(() => {
          this.navigateToSource();// should stop here when on a form page// will navigate back to listview 
          return newTR;
        });

      });
    }
  }

  /**
   * Emails the StaffCC if the TR has just been completed
   * 
   * @private
   * @param {TR} tr The TR we saved.
   * @param {string} originalStatus The status of the TR Prior to us saving it,.
   * @returns {Promise<any>} 
   * 
   * @memberof TrFormWebPart
   */
  private emailStaffCC(tr: TR, originalStatus: string): Promise<any> {

    return new Promise((resolve, reject) => {
      if (!this.properties.enableEmail || tr.TRStatus != "Completed" || originalStatus === "Completed" || tr.StaffCC === null || tr.StaffCC.length === 0) {
        console.log("staffcc emails wil lnot be processed");
        resolve(null);
        return;
      }
      let promises: Array<Promise<any>> = [];
      let editFormUrl = this.properties.editFormUrlFormat.replace("{1}", tr.Id.toString());
      editFormUrl = editFormUrl.split("{0}").join(this.context.pageContext.web.absoluteUrl); //split&join to replace all
      let displayFormUrl = this.properties.displayFormUrlFormat.replace("{1}", tr.Id.toString());
      displayFormUrl = displayFormUrl.split("{0}").join(this.context.pageContext.web.absoluteUrl); //split&join to replace all
      console.log("fetching email text in emailStaffCC");
      var y = pnp.sp.web.lists.getByTitle(this.properties.setupListName).items.getAs<SetupItem[]>().then((setupItems) => {
        console.log("fetched email text in emailStaffCC, extracting text");
        let subject: string = _.find(setupItems, (si: SetupItem) => { return si.Title === "StaffCC Email Subject"; }).PlainText
          .replace("~technicalRequestNumber", tr.Title)
          .replace("~technicalRequestEditUrl", editFormUrl)
          .replace("~technicalRequestDisplayUrl", displayFormUrl);
        let body: string = _.find(setupItems, (si: SetupItem) => { return si.Title === "StaffCC Email Body"; }).RichText
          .replace("~technicalRequestNumber", tr.Title)
          .replace("~technicalRequestEditUrl", editFormUrl)
          .replace("~technicalRequestDisplayUrl", displayFormUrl);
        console.log("extracted text in emailStaffCC, looping users");
        for (let staffCC of tr.StaffCC) {
          console.log("in emailStaffCC, fetching user " + staffCC);
          //*******          TODO , O a;ready have the email address in the persona


          let promise = pnp.sp.web.getUserById(parseInt(staffCC.id)).get().then((user) => {
            console.log("in emailStaffCC, fetched user " + staffCC);
            let emailProperties: EmailProperties = {
              From: this.context.pageContext.user.email,
              To: [user.Email],
              Subject: subject,
              Body: body
            };
            console.log("in emailStaffCC, emailing user " + user.Email);
            return pnp.sp.utility.sendEmail(emailProperties)
              .then((x) => {
                console.log("email sent to " + emailProperties.To);
              })
              .catch((error) => {
                console.log(error);
              });

          }).catch((error) => {
            console.log("Error Fetching user with id " + staffCC);
          });
          promises.push(promise);
        }
        Promise.all(promises).then((x) => {
          resolve();
        });
      });
    });
  }

  /**
   * Sends notifications to any new assignees.
   * 
   * @private
   * @param {TR} tr The TR we just saved,
   * @param {Array<number>} orginalAssignees The list of assignees prior to us saving the TR
   * @returns {Promise<any>} 
   * 
   * @memberof TrFormWebPart
   */
  private emailNewAssignees(tr: TR, orginalAssignees: Array<number>): Promise<any> {
    return new Promise((resolve, reject) => {
      if (!this.properties.enableEmail) {
        resolve(null);
        return;
      }
      if (tr.TRAssignedToId === null || tr.TRAssignedToId.length === 0) {
        resolve(null);
        return;
      }

      let promises: Array<Promise<any>> = [];
      let currentAssignees: Array<number> = tr.TRAssignedToId;
      let editFormUrl = this.properties.editFormUrlFormat
        .split("{1}").join(tr.Id.toString())
        .split("{0}").join(this.context.pageContext.web.absoluteUrl);
      let displayFormUrl = this.properties.displayFormUrlFormat
        .split("{1}").join(tr.Id.toString())
        .split("{0}").join(this.context.pageContext.web.absoluteUrl);
      console.log("fetching email text in emailNewAssignees");
      var x = pnp.sp.web.lists.getByTitle(this.properties.setupListName).items.getAs<SetupItem[]>().then((setupItems) => {
        console.log("fetched email text in emailNewAssignees, extracting it now");
        let subject: string = _.find(setupItems, (si: SetupItem) => { return si.Title === "Assignee Email Subject"; }).PlainText
          .replace("~technicalRequestNumber", tr.Title)
          .replace("~technicalRequestEditUrl", editFormUrl)
          .replace("~technicalRequestDisplayUrl", displayFormUrl);
        let body: string = _.find(setupItems, (si: SetupItem) => { return si.Title === "Assignee Email Body"; }).RichText
          .replace("~technicalRequestNumber", tr.Title)
          .replace("~technicalRequestEditUrl", editFormUrl)
          .replace("~technicalRequestDisplayUrl", displayFormUrl);
        console.log("extracted email text in emailNewAssignees,looping assignees");
        console.log("cuurnt assignees are:" + currentAssignees);

        for (let assignee of currentAssignees) {
          if (orginalAssignees === null || orginalAssignees.indexOf(assignee) === -1) {
            // send email
            console.log("fetchin assignee #" + assignee);
            let promise = pnp.sp.web.getUserById(assignee).get().then((user) => {
              console.log("fetche assignee #" + assignee);
              let emailProperties: EmailProperties = {
                From: this.context.pageContext.user.email,
                To: [user.Email],
                Subject: subject,
                Body: body
              };
              console.log("dending email to assignee assignee #" + assignee + "  " + user.Email);
              return pnp.sp.utility.sendEmail(emailProperties)
                .then((resp) => {
                  console.log("Assignee email sent to " + emailProperties.To);
                })
                .catch((error) => {
                  console.log(error);
                  reject(error);
                });

            }).catch((error) => {
              console.log("Error Fetching user with id " + assignee);
            });
            promises.push(promise);
          }
          else {
            console.log("asignee is not new");
          }
        }
        Promise.all(promises).then((resp) => {
          resolve();
        });
      }).catch((error) => {

        console.log(error);
      });


    });
  }

  /**
   * Closes the app by navigating to teh source
   * 
   * @private
   * 
   * @memberof TrFormWebPart
   */
  private cancel(): void {
    this.navigateToSource();
  }

  /**
   * Uploads a file to the TR DOcument library an associates it with the specified TR
   * 
   * @private
   * @param {any} file The file to upload
   * @param {any} trId  The ID of the TR to associate the file with
   * @returns {Promise<any>} 
   * 
   * @memberof TrFormWebPart
   */
  private uploadFile(file, trId): Promise<any> {
    if (file.size <= 10485760) {
      // small upload
      return pnp.sp.web.lists.getByTitle(this.properties.trDocumentsListName).rootFolder.files.add(file.name, file, true)
        .then((results) => {
          //return pnp.sp.web.getFileByServerRelativeUrl(results.data.ServerRelativeUrl).getItem<{ Id: number, Title: string, Modified: Date }>("Id", "Title", "Modified").then((item) => {
          return pnp.sp.web.getFileByServerRelativeUrl(results.data.ServerRelativeUrl).getItem().then((item) => {

            const itemID = parseInt(item["Id"]);
            return pnp.sp.web.lists.getByTitle(this.properties.trDocumentsListName).items.getById(itemID).update({ "TRId": trId, Title: file.name })
              .then((response) => {

                return;
              }).catch((error) => {

              });
          }).catch((error) => {
            console.log(error);
          });

        }).catch((error) => {
          console.log(error);
        });
    } else {
      // large upload// not tested yet
      alert("large file support  not impletemented");

      return pnp.sp.web.lists.getByTitle(this.properties.trDocumentsListName).rootFolder.files
        .addChunked(file.name, file, data => {
          console.log({ data: data, message: "progress" });
        }, true)
        .then((results) => {
          console.log("done!");
        })
        .catch((error) => {

          console.log(error);
        });
    }
  }
}

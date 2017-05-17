import * as React from 'react';
import * as ReactDom from 'react-dom';
import pnp from "sp-pnp-js";
import { SearchQuery, SearchResults, SortDirection, EmailProperties } from "sp-pnp-js";
import { Version, UrlQueryParameterCollection } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration, PropertyPaneDropdown, PropertyPaneTextField, PropertyPaneCheckbox
} from '@microsoft/sp-webpart-base';
import {
  Environment,
  EnvironmentType
} from '@microsoft/sp-core-library';
import { SetupItem, Test, PropertyTest, Pigment, TR, WorkType, ApplicationType, EndUse, modes, User, Customer } from "./dataModel";
import * as strings from 'trFormStrings';
import * as _ from 'lodash';
import TrForm from './components/TrForm';
import { ITrFormProps } from './components/ITrFormProps';
import { ITrFormWebPartProps } from './ITrFormWebPartProps';
export default class TrFormWebPart extends BaseClientSideWebPart<ITrFormWebPartProps> {
  private tr: TR;
  private childTRs: Array<TR>;



  private reactElement: React.ReactElement<ITrFormProps>;
  private trContentTypeID: string;
  public onInit(): Promise<void> {

    return super.onInit().then(_ => {

      pnp.setup({
        spfxContext: this.context
      });
      //  this.loadData();
    });
  }
  private moveFieldsToTR(tr: TR, item: any) {
    tr.Id = item.Id;
    tr.ActualCompletionDate = item.ActualCompletionDate;
    tr.ApplicationTypeId = item.ApplicationTypeId;
    tr.ActualStartDate = item.ActualStartDate;
    tr.CER = item.CER;
    tr.CustomerId = item.CustomerId;

    tr.RequiredDate = item.RequiredDate;
    tr.EstManHours = item.EstimatedHours;
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
    tr.RequestTitle = item.ReqquestTitle;
    tr.Formulae = item.Formulae;
    tr.Description = item.Description;
    tr.Summary = item.Summary;
    tr.TestingParameters = item.TestingParameters;
    tr.ParentTRId = item.ParentTRId;
    if (item.ParentTR) {
      tr.ParentTR = item.ParentTR.Title;
    }
    tr.TRAssignedToId = item.TRAssignedToId;
    tr.StaffCCId = item.StaffCCId;
    tr.PigmentsId = item.PigmentsId;
    tr.TestsId = item.TestsId;
  }
  public fetchTR(id: number): Promise<TR> {
    let fields = "*,ParentTR/Title,Requestor/Title";
    return pnp.sp.web.lists.getByTitle(this.properties.technicalRequestListName).items.getById(id).expand("ParentTR,Requestor").select(fields).get()
      .then((item) => {
        let tr = new TR();
        this.moveFieldsToTR(tr, item);
        return tr;
      });
  }
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
  // An accesser indicating whether or not the current page is in design mode.
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
  public render(): void {
    // hide the ribbon
    //if (!this.inDesignMode())
    if (document.getElementById("s4-ribbonrow")) {
      document.getElementById("s4-ribbonrow").style.display = "none";
    }

    let formProps: ITrFormProps = {
      customers: [],
      subTRs: [],
      techSpecs: [],
      requestors: [],
      cancel: this.cancel.bind(this),
      mode: this.properties.mode,
      TRsearch: this.TRsearch.bind(this),
      workTypes: [],
      applicationTypes: [],
      endUses: [],
      tr: new TR(),
      save: this.save.bind(this),
      fetchChildTr: this.fetchChildTR.bind(this),
      fetchTR: this.fetchTR.bind(this),
      pigments: [],
      tests: [],
      propertyTests: []
    };
    let batch = pnp.sp.createBatch();
    // get the Technincal Request content type so we can use it later in searches
    // pnp.sp.web.contentTypes.inBatch(batch).get()
    //   .then((contentTypes) => {

    //     const trContentTyoe = _.find(contentTypes, (contentType) => { return contentType["Name"] === "TechnicalRequest"; });
    //     this.trContentTypeID = trContentTyoe["Id"]["StringValue"];
    //   })
    //   .catch((error) => {
    //     console.log("ERROR, An error occured fetching content types'");
    //     console.log(error.message);

    //   });



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
    pnp.sp.web.lists.getByTitle(this.properties.workTypeListName).items.inBatch(batch).get()
      .then((items) => {
        formProps.workTypes = _.map(items, (item) => {
          return new WorkType(item["Id"], item["Title"]);
        });

      })
      .catch((error) => {
        console.log("ERROR, An error occured fetching 'Work Types' from list named " + this.properties.workTypeListName);
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
        let fields = "*,ParentTR/Title,Requestor/Title,Customer/Title,TRAssignedTo/Title,TRAssignedTo/Id";
        let expands = "ParentTR,Requestor,Customer,TRAssignedTo";
        // get the requested tr
        pnp.sp.web.lists.getByTitle(this.properties.technicalRequestListName).items.getById(id).expand(expands).select(fields).inBatch(batch).get()

          .then((item) => {
            formProps.tr = new TR();
            this.moveFieldsToTR(formProps.tr, item);
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
          })
          .catch((error) => {
            console.log("ERROR, An error occured fetching the listitem  from list named " + this.properties.technicalRequestListName);
            console.log(error.message);

          });
        // get the Child trs

        pnp.sp.web.lists.getByTitle(this.properties.technicalRequestListName).items.filter("ParentTRId eq " + id).expand(expands).select(fields).inBatch(batch).get()

          .then((items) => {
            // this may resilve befor we get the mainn tr, so jyst stash them away for now.
            for (const item of items) {
              let childtr: TR = new TR();
              this.moveFieldsToTR(childtr, item);
              formProps.subTRs.push(childtr);
            }
          })
          .catch((error) => {
            console.log("ERROR, An error occured fetching child trs  from list named " + this.properties.technicalRequestListName);
            console.log(error.message);

          });
      }
      else {
        console.log("ERROR, Id not specified with Display or Edit form");
      }
    }
    else {
      formProps.tr.Site = this.properties.defaultSite;
      pnp.sp.web.lists.getByTitle(this.properties.nextNumbersListName).items.select("Id,NextNumber").filter("CounterName eq 'RequestId'").orderBy("Title").top(5000).inBatch(batch).get()// get the lookup info
        .then((items) => {
          if (items.length != 1) {
            console.log("multiple next numbers found");
          }
          else {
            let nextNumber: number = items[0]["NextNumber"];
            nextNumber++;
            formProps.tr.Title = this.properties.defaultSite + nextNumber;

            pnp.sp.web.lists.getByTitle(this.properties.nextNumbersListName).items.getById(items[0].Id)
              .update({ "NextNumber": nextNumber }).then((results) => {
                console.log("next number not increment to " + nextNumber);
              }).catch((err) => {
                alert("next number not incremented-- please try again");
              });
          }
        }).catch((err) => {
          debugger;
          console.log("next number not increment to");
        });
    }


    batch.execute().then((value) => {// execute the batch to get the item being edited and info REQUIRED for initial display
      this.reactElement = React.createElement(TrForm, formProps);
      var formComponent: TrForm = ReactDom.render(this.reactElement, this.domElement) as TrForm;//render the component
      if (Environment.type === EnvironmentType.ClassicSharePoint) {
const buttons: NodeListOf<HTMLButtonElement> = this.domElement.getElementsByTagName('button');
if (buttons && buttons.length) {
for (let i: number = 0; i < buttons.length; i++) {
if (buttons[i]) {
/* tslint:disable */
// Disable the button onclick postback
buttons[i].onclick = function() { return false; };
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
      pnp.sp.web.lists.getByTitle(this.properties.partyListName).items.select(customerFields).filter("Active eq -1 ").orderBy("Title").top(5000).inBatch(batch2).get()// get the lookup info
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

      let pigmentFields = "Id,Title,KMPigmentType,Manufacturer/Title";
      let pigmentExpands = "Manufacturer";
      pnp.sp.web.lists.getByTitle(this.properties.pigmentListName).items.select(pigmentFields).expand(pigmentExpands).top(5000).inBatch(batch2).get()// get the lookup info
        .then((items) => {
          formProps.pigments = _.map(items, (item) => {
            let p: Pigment = new Pigment(item["Id"], item["Title"], item["KMPigmentType"]);
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
        // this.delay(5000).then(() => {
        formComponent.props.customers = formProps.customers;
        formComponent.props.pigments = formProps.pigments;
        formComponent.props.tests = formProps.tests;
        formComponent.props.propertyTests = formProps.propertyTests;
        formComponent.props.techSpecs = formProps.techSpecs;
        formComponent.props.requestors = formProps.requestors;



        formComponent.forceUpdate();
        // });


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
                PropertyPaneCheckbox('enableEmail', {
                  text: "Enable sending emails to assignees and staff cc",
                })
              ]
            }
          ]
        }
      ]
    };
  }
  public TRsearch(searchText: string): Promise<TR[]> {

    //let queryText = "{0} Path:{1}* ContentTypeId:{2}*";
    let queryText = "{0} Path:{1}*";
    queryText = queryText
      .replace("{0}", searchText)
      .replace("{1}", this.properties.searchPath);
    //.replace("{2}", this.trContentTypeID);
    let sq: SearchQuery = {
      Querytext: queryText,
      RowLimit: 50,
      SelectProperties: ["Title", "ListItemID", "CEROWSTEXT", "CustomerOWSTEXT", "SiteOWSTEXT"]
      ///SortList: [{ Property: "PreferredName", Direction: SortDirection.Ascending }] arghhh-- not sortable
      // SelectProperties: ["*"]
    };
    console.log(sq);

    return pnp.sp.search(sq).then((results: SearchResults) => {
      let returnValue: Array<TR> = [];
      for (let sr of results.PrimarySearchResults) {
        const temp = sr as any;
        let tr: TR = new TR();
        tr.Id = temp.ListItemID;
        tr.Title = temp.Title;
        tr.CustomerId = temp.CustomerOWSTEXT;
        tr.Site = temp.SiteOWSTEXT;
        tr.CER = temp.CEROWSTEXT;
        returnValue.push(tr);
      }


      return _.sortBy(returnValue, "Title");
    });


  }


  private navigateToSource() {
    debugger;
    let queryParameters = new UrlQueryParameterCollection(window.location.href);
    let encodedSource = queryParameters.getValue("Source");
    if (encodedSource) {
      let source = decodeURIComponent(encodedSource);
      console.log("source uis " + source);
      window.location.href = source;
    }
  }
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


    let StaffCCId = (copy.StaffCCId) ? copy.StaffCCId : [];
    delete copy.StaffCCId;
    copy["StaffCCId"] = {};
    copy["StaffCCId"]["results"] = StaffCCId;

    let PigmentsId = (copy.PigmentsId) ? copy.PigmentsId : [];
    delete copy.PigmentsId;
    copy["PigmentsId"] = {};
    copy["PigmentsId"]["results"] = PigmentsId;


    let TestsId = (copy.TestsId) ? copy.TestsId : [];
    delete copy.TestsId;
    copy["TestsId"] = {};
    copy["TestsId"]["results"] = TestsId;

    if (copy.Id !== null) {
      return pnp.sp.web.lists.getByTitle(this.properties.technicalRequestListName).items.getById(tr.Id).update(copy).then((item) => {
        let newAssigneesPromise = this.emailNewAssignees(tr, orginalAssignees);
        var staffccPromise = this.emailStaffCC(tr, originalStatus);
        Promise.all([newAssigneesPromise, staffccPromise]).then((a) => {
          let x = newAssigneesPromise;
          let y = staffccPromise;
          this.navigateToSource();// should stop here when on a form page  
          return tr;
        })
      });
    }
    else {
      delete copy.Id;
      return pnp.sp.web.lists.getByTitle(this.properties.technicalRequestListName).items.add(copy).then((item) => {
        let newTR = new TR();
        this.moveFieldsToTR(newTR, item.data);
        var newAssigneesPromise = this.emailNewAssignees(newTR, orginalAssignees);
        var staffccPromise = this.emailStaffCC(newTR, originalStatus);
        Promise.all([newAssigneesPromise, staffccPromise]).then(() => {
          this.navigateToSource();// should stop here when on a form page  
          return newTR;
        })

      });

    }

  }
  private emailStaffCC(tr: TR, originalStatus: string): Promise<any> {
    return new Promise((resolve, reject) => {
      debugger;
      if (!this.properties.enableEmail || tr.TRStatus != "Completed" || originalStatus === "Completed" || tr.StaffCCId === null || tr.StaffCCId.length === 0) {
        debugger;
        resolve(null);
        return;
      }
      let promises: Array<Promise<any>> = [];
      let editFormUrl = this.properties.editFormUrlFormat.replace("{1}", tr.Id.toString());
      let displayFormUrl = this.properties.displayFormUrlFormat.replace("{1}", tr.Id.toString());
      var y = pnp.sp.web.lists.getByTitle(this.properties.setupListName).items.getAs<SetupItem[]>().then((setupItems) => {
        let subject: string = _.find(setupItems, (si: SetupItem) => { return si.Title === "StaffCC Email Subject"; }).PlainText
          .replace("~technicalRequestNumber", tr.Title)
          .replace("~technicalRequestEditUrl", editFormUrl)
          .replace("~technicalRequestDisplayUrl", displayFormUrl);
        let body: string = _.find(setupItems, (si: SetupItem) => { return si.Title === "StaffCC Email Body"; }).RichText
          .replace("~technicalRequestNumber", tr.Title)
          .replace("~technicalRequestEditUrl", editFormUrl)
          .replace("~technicalRequestDisplayUrl", displayFormUrl);
       for (let staffCC of tr.StaffCCId) {
          let promise = pnp.sp.web.getUserById(staffCC).get().then((user) => {
            let emailProperties: EmailProperties = {
              From: this.context.pageContext.user.email,
              To: [user.Email],
              Subject: subject,
              Body: body
            };
            debugger;
            return pnp.sp.utility.sendEmail(emailProperties)
              .then((x) => {
                console.log("email sent to " + emailProperties.To);
              })
              .catch((error) => {
                debugger;
                console.log(error);
              });

          }).catch((error) => {
            console.log("Error Fetching user with id " + staffCC);
          });
          promises.push(promise);
          debugger;
        }
        Promise.all(promises).then((x) => {
          debugger;
          resolve();
        });
      });


    });
  }
  private emailNewAssignees(tr: TR, orginalAssignees: Array<number>): Promise<any> {
    return new Promise((resolve, reject) => {

      if (!this.properties.enableEmail) {

        resolve(null);
        return;
      }
      let promises: Array<Promise<any>> = [];
      let currentAssignees: Array<number> = tr.TRAssignedToId;
      let editFormUrl = this.properties.editFormUrlFormat.replace("{1}", tr.Id.toString());
      let displayFormUrl = this.properties.displayFormUrlFormat.replace("{1}", tr.Id.toString());
      var x = pnp.sp.web.lists.getByTitle(this.properties.setupListName).items.getAs<SetupItem[]>().then((setupItems) => {
        let subject: string = _.find(setupItems, (si: SetupItem) => { return si.Title === "Assignee Email Subject"; }).PlainText
          .replace("~technicalRequestNumber", tr.Title)
          .replace("~technicalRequestEditUrl", editFormUrl)
          .replace("~technicalRequestDisplayUrl", displayFormUrl);
        let body: string = _.find(setupItems, (si: SetupItem) => { return si.Title === "Assignee Email Body"; }).RichText
          .replace("~technicalRequestNumber", tr.Title)
          .replace("~technicalRequestEditUrl", editFormUrl)
          .replace("~technicalRequestDisplayUrl", displayFormUrl);
        for (let assignee of currentAssignees) {
          if (orginalAssignees === null || orginalAssignees.indexOf(assignee) === -1) {
            // send email
            let promise = pnp.sp.web.getUserById(assignee).get().then((user) => {
              let emailProperties: EmailProperties = {
                From: this.context.pageContext.user.email,
                To: [user.Email],
                Subject: subject,
                Body: body
              };
             return pnp.sp.utility.sendEmail(emailProperties)
                .then((x) => {
                  console.log("Assignee email sent to "+emailProperties.To);
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
        }
        Promise.all(promises).then((x) => {
          resolve();
        });
      }).catch((error) => {

        console.log(error);
      });


    });
  }
  private cancel(): void {
    this.navigateToSource();

  }
}


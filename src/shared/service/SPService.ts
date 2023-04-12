import { WebPartContext } from "@microsoft/sp-webpart-base";
import { sp } from '@pnp/sp/presets/all';

export class SPService {
    constructor(private context: WebPartContext) {
        sp.setup({
            spfxContext: this.context
        });
    }

    public async getListItems(selectedList: string, selectedFields: any[]) {
        try {
            let selectQuery: any[] = ['Id'];
            let expandQuery: any[] = [];
            let listItems = [];
            let items: any;
            for (var i = 0; i < selectedFields.length; i++) {
                switch (selectedFields[i].fieldType) {
                    case 'SP.FieldUser':
                        selectQuery.push(`${selectedFields[i].key}/Title,${selectedFields[i].key}/EMail,${selectedFields[i].key}/Name`);
                        expandQuery.push(selectedFields[i].key);
                        break;
                    case 'SP.FieldLookup':
                        selectQuery.push(`${selectedFields[i].key}/Title`);
                        expandQuery.push(selectedFields[i].key);
                        break;
                    case 'SP.Field':
                        selectQuery.push('Attachments,AttachmentFiles');
                        expandQuery.push('AttachmentFiles');
                        break;
                    default:
                        selectQuery.push(selectedFields[i].key);
                        break;
                }
            }
            items = await sp.web.lists.getById(selectedList).items
                .select(selectQuery.join())
                .expand(expandQuery.join())
                .top(4999)
                .getPaged();
            listItems = items.results;
            while (items.hasNext) {
                items = await items.getNext();
                listItems = [...listItems, ...items.results];
            }
            return listItems;
        } catch (err) {
            Promise.reject(err);
        }
    }

    //MyFunctions

    public async UpdateData(itemId,selectedList: string,Status:string,Mydate:string)
    {

    
    const Item = await sp.web.lists.getById(selectedList).items.getById(itemId).update({
    ApproverStatus:Status,
    Test:Mydate
       
      });
    
    console.log(itemId);
    };

    public async getListItemsBasedonApproverID(selectedList: string, selectedFields: any[],ApproverID:string) {
        try {
            let mystatus="Pending";
            let selectQuery: any[] = ['Id'];
            let expandQuery: any[] = [];
            let listItems = [];
            let items: any;
            for (var i = 0; i < selectedFields.length; i++) {
                switch (selectedFields[i].fieldType) {
                    case 'SP.FieldUser':
                        selectQuery.push(`${selectedFields[i].key}/Title,${selectedFields[i].key}/EMail,${selectedFields[i].key}/Name`);
                        expandQuery.push(selectedFields[i].key);
                        break;
                    case 'SP.FieldLookup':
                        selectQuery.push(`${selectedFields[i].key}/Title`);
                        expandQuery.push(selectedFields[i].key);
                        break;
                    case 'SP.Field':
                        selectQuery.push('Attachments,AttachmentFiles');
                        expandQuery.push('AttachmentFiles');
                        break;
                    default:
                        selectQuery.push(selectedFields[i].key);
                        break;
                }
            }
            items = await sp.web.lists.getById(selectedList).items
                .select(selectQuery.join())
                .filter("ApproverID eq '" +ApproverID+"' and ApproverStatus eq '" +mystatus+"'")
                .expand(expandQuery.join())
                .top(4999)
                .getPaged();
            listItems = items.results;
            while (items.hasNext) {
                items = await items.getNext();
                listItems = [...listItems, ...items.results];
            }
            return listItems;
        } catch (err) {
            Promise.reject(err);
        }
    }

    public async updateMainList(itemId,selectedList: string,Status:string,ApproverRecID:string,ApproverRecName:string,MyComments:string)
    {
        let Myval='Completed';
        let Varmyval = await sp.web.lists.getById(selectedList).items.getById(itemId).update({
        ApproverStatus:Status,
        ApproverID:"",
        ApproverName:"",
        ApproverID_Old:ApproverRecID,
        ApproverName_Old:ApproverRecName,
        Comments:MyComments

      });
            
    console.log(itemId);

    }
    
    public async InserttoBulkList(MyTitle:string,MyApproverID:string,MyApproverName:string,MyEmpname:string,
    MyEmpID:string,MyReqQuarter:string,MyReqYear:string,MyReviewdate:Date,
    MyReqStatus:string,MyComments:string,MyListName:String)

    {

        let Varmyval= await sp.web.lists.getByTitle("BulkApprovalDetatilsList").items.add({
        Title:MyTitle,
        ApproverID:MyApproverID,
        ApproverName:MyApproverName,
        EmpName:MyEmpname,
        EmpId:MyEmpID,
        MyQuarter:MyReqQuarter,
        MyYear:MyReqYear,
        WF_ReviewDueDate:MyReviewdate,
        MyStatus:MyReqStatus,
        Comments:MyComments,
        NameofList:MyListName
                   
        });

    }

    public async InserttoBulkLists(MyTitle:string,MyApproverID:string,MyApproverName:string,MyEmpname:string,
        MyEmpID:string,MyReqQuarter:string,MyReqYear:string,MyReviewdate:Date,
        MyReqStatus:string,MyComments:string,MyListName:String)
    
        {
    
            let Varmyval= await sp.web.lists.getByTitle("BulkApprovalDetatilsList").items.add({
                ReqIDS:MyTitle,
                ApproverID:MyApproverID,
                ApproverName:MyApproverName,
                EmpName:MyEmpname,
                EmpId:MyEmpID,
                MyQuarter:MyReqQuarter,
                MyYear:MyReqYear,
                WF_ReviewDueDate:MyReviewdate,
                MyStatus:MyReqStatus,
                Comments:MyComments,
                NameofList:MyListName
            
            });
    
        }


    //End

    public async getFields(selectedList: string): Promise<any> {
        try {
            const allFields: any[] = await sp.web.lists
                .getById(selectedList)
                .fields
                .filter("Hidden eq false and ReadOnlyField eq false and Title ne 'Content Type' and Title ne 'Attachments'")
                .get();
            return allFields;
        }
        catch (err) {
            Promise.reject(err);
        }
    }

    public async getUserProfileUrl(loginName: string) {
        try {
            const properties = await sp.profiles.getPropertiesFor(loginName);
            const profileUrl = properties['PictureUrl'];
            return profileUrl;
        }
        catch (err) {
            Promise.reject(err);
        }
    }


    public async getItemIDs(selectedList: string, data:string,ApproverID:string): Promise<any> {
       
        let mystatus="pending";
        let filtercondition: any = "(Title eq '" + data + "' and ApproverStatus eq '" +mystatus+"' and ApproverID eq '"+ApproverID+"' )";
        const allItems: any[] = await sp.web.lists.getByTitle(selectedList).items.filter(filtercondition).getAll();
        return allItems;


    }


}

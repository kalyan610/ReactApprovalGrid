import * as React from 'react';
import styles from './ReactDatatable.module.scss';


import { IReactDatatableProps } from './IReactDatatableProps';
import { IReactDatatableState } from './IReactDatatableState';
import * as strings from 'ReactDatatableWebPartStrings';
import { SPService } from '../../../shared/service/SPService';
import { Placeholder } from "@pnp/spfx-controls-react/lib/Placeholder";
import { DisplayMode } from '@microsoft/sp-core-library';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Grid } from '@material-ui/core';
import { Link, PrimaryButton, Stack, Text,StackItem,IStackTokens,IStackStyles} from 'office-ui-fabric-react';
import { WebPartTitle } from "@pnp/spfx-controls-react/lib/WebPartTitle";
import { ExportListItemsToCSV } from '../../../shared/common/ExportListItemsToCSV/ExportListItemsToCSV';
import { ExportListItemsToPDF } from '../../../shared/common/ExportListItemsToPDF/ExportListItemsToPDF';
import { Pagination } from '../../../shared/common/Pagination/Pagination';
import { RenderImageOrLink } from '../../../shared/common/RenderImageOrLink/RenderImageOrLink';
import { DetailsList, DetailsListLayoutMode, DetailsRow, IDetailsRowStyles, IDetailsListProps, IColumn, MessageBar, SelectionMode,Selection } from 'office-ui-fabric-react';
import { pdfCellFormatter } from '../../../shared/common/ExportListItemsToPDF/ExportListItemsToPDFFormatter';
import { csvCellFormatter } from '../../../shared/common/ExportListItemsToCSV/ExportListItemsToCSVFormatter';
import { IPropertyPaneDropdownOption } from '@microsoft/sp-property-pane';
import { RenderProfilePicture } from '../../../shared/common/RenderProfilePicture/RenderProfilePicture';
import { Dialog, DialogFooter, DialogType } from 'office-ui-fabric-react/lib/Dialog'; 

const sectionStackTokens: IStackTokens = { childrenGap: 5 };
const sectionStackTokens1: IStackTokens = { childrenGap: 5 };
const stackTokens = { childrenGap: 80 };
const stackStyles: Partial<IStackStyles> = { root: { padding: 10 } };
const stackButtonStyles: Partial<IStackStyles> = { root: { Width: 20 } };

let GlobalMyArray=[];

let GlobalBulkIDs=[];

let AllListItems=[];

export default class ReactDatatable extends React.Component<IReactDatatableProps, IReactDatatableState> {

  private _services: SPService = null;
   
  private _selection:Selection;
 

  constructor(props: IReactDatatableProps) {
    super(props);
    this.state = {
      listItems: [],
      columns: [],
      page: 1,
      searchText: '',
      rowsPerPage: 10,
      sortingFields: '',
      sortDirection: 'asc',
      contentType: '',
      pageOfItems: [],
      SelectionDetails:'',
      myApproverId:'',
      openbox: true,
      EmpID:"",
      EmpName:"",
      SelectedMyArary:[],
      Comments:""
            
    };

        
    

    this.loadItems();
    this._services = new SPService(this.props.context);
    this._onConfigure = this._onConfigure.bind(this);
    this.getSelectedListItems = this.getSelectedListItems.bind(this);
    this._selection = new Selection({ });
    
      
    

  }

  public componentDidMount() {

   
    this.getSelectedListItems();

    // let { ToggleHide } = this.props;

    // if(!ToggleHide)
    // {

    //   this.getloginuser();
    // }

    if(this.props.AdminView==false)
    {
    this.getloginuser();
    }
  }

  //My functions

private  Approve()
  {

  // const current = new Date();

  // let testdate = new Date(current);
  // testdate.setDate(testdate.getDate()-1);
  // let x=testdate.toDateString();

 console.log(this._selection.getSelection());

 let myApproverArray=[];

 myApproverArray=this._selection.getSelection();

 console.log(myApproverArray.length);

 if(myApproverArray.length)
 {

 for(let count=0;count<myApproverArray.length;count++)
 {

  this._services.UpdateData(myApproverArray[count].id,this.props.list,"Approved").then(function (data)
  {

    if(count==myApproverArray.length-1)
    {
    alert('Approved successfully');
    }
    window.location.reload();
    
    

  });

 }



}
else
{
 
alert('Please select atleast one record to approve');
}

}

private Remove()
{

  const current = new Date();


  
  console.log(this._selection.getSelection());

 let myApproverArray=[];

 myApproverArray=this._selection.getSelection();

 console.log(myApproverArray.length);

 if(myApproverArray.length)
 {


 for(let count=0;count<myApproverArray.length;count++)
 {
  
  this._services.UpdateData(myApproverArray[count].id,this.props.list,"Removed").then(function (data)
  {

    if(count==myApproverArray.length-1)
    {
    alert('Removed successfully');
    }
    window.location.reload();
    
  });

 }

 
}
else
{
 
alert('Please select atleast one record to approve');
}

}

public checkduplicateempvalues(arr)
				{

          
				
				var xapp=arr[0];

        this.setState({EmpID:xapp['Title']});
        this.setState({EmpName: xapp['EmpName'] });

			  for (var i=0;i<arr.length;i++){
				if(xapp['Title']!=arr[i]['Title'])
        { 
        alert('The selected records must contain the same employee to perform a bulk approver change');
				return false;
				}
				}

        return true;
								
}

public changerecord()
{

  console.log(this._selection.getSelection());

  let myApproverArray=[];
 
  myApproverArray=this._selection.getSelection();
 
  GlobalMyArray=myApproverArray;
  

if(myApproverArray.length)
  {

   if(this.checkduplicateempvalues(myApproverArray))
    {

      
      this.setState({openbox: false });

    }

  }

  else
  {

    alert('Please select atleast one record to change');
  }
}

public saveChange()
{

  
  //Update record to in Main List
  
  let myRecordIDS="";
  
 console.log(GlobalMyArray.length);

 if(GlobalMyArray.length)
 {

 
 for(let count=0;count<GlobalMyArray.length;count++)
 {
  myRecordIDS+=GlobalMyArray[count].id+",";


  // this._services.updateMainList(GlobalMyArray[count].id,"SAR_PeopleSoftFinance","BulkChange",GlobalMyArray[count].ApproverID,GlobalMyArray[count].ApproverName,this.state.Comments).then(function (data)
  // {
       
  // });

  //this._services.updateMainList(GlobalMyArray[count].id,this.props.list,"BulkChange",GlobalMyArray[count].ApproverID,GlobalMyArray[count].ApproverName,this.state.Comments);
  
 }

 //End

 //Isert to Bulk List

 let myRecordIDSItems =myRecordIDS.slice(0, -1);

console.log(myRecordIDSItems)

  let testdate = new Date(GlobalMyArray[0].ReviewDueDate); 

  testdate.setDate(testdate.getDate()+1);

  let Finaldate=new Date(testdate);
   console.log('kalyan testing');

this._services.InserttoBulkList(myRecordIDSItems,GlobalMyArray[0].ApproverID,GlobalMyArray[0].ApproverName,GlobalMyArray[0].EmpName,
 GlobalMyArray[0].Title,GlobalMyArray[0].WF_Quarter,GlobalMyArray[0].WF_Year,Finaldate,"BulkChange",this.state.Comments,this.props.title).then(function (data)
 {

 
  alert('Changed successfully');

  window.location.reload();
      
 });
 

 this.setState({openbox: true });
 
 

//End

}


else
{
 
alert('Please select atleast one record to approve');
}
 //End

}

public async savechangetest()
{

var FinalArray=[];

let myRecordIDSItems="";

let myIDSarray=[];

let MyIdElements;

FinalArray=this.removewithfilter(GlobalMyArray)

//let listItemIDS = await this._services.getItemIDs(FinalArray[0]['Title']);
let reqApproverIDSIDs=this.getParam("SID");



let listItemIDS = await this._services.getItemIDs1(this.props.list,FinalArray[0]['Title'],reqApproverIDSIDs);



let testdate = new Date(GlobalMyArray[0].ReviewDueDate); 

testdate.setDate(testdate.getDate()+1);

let Finaldate=new Date(testdate);

 

for(var count=0;count<listItemIDS.length;count++)
{


  myRecordIDSItems+=listItemIDS[count].ID+",";

 await this._services.updateMainList(listItemIDS[count].ID,this.props.list,"BulkChange",listItemIDS[count].ApproverID,listItemIDS[count].ApproverName,this.state.Comments).then(function (data)
  {
       
 });

//this._services.updateMainList(listItemIDS[count].ID,"SAR_PeopleSoftFinance","BulkChange",listItemIDS[count].ApproverID,listItemIDS[count].ApproverName,this.state.Comments);

}
//Bulk Save

let FinalmyRecordIDSItems =myRecordIDSItems.slice(0, -1);

this._services.InserttoBulkLists(FinalmyRecordIDSItems,GlobalMyArray[0].ApproverID,GlobalMyArray[0].ApproverName,GlobalMyArray[0].EmpName,
  GlobalMyArray[0].Title,GlobalMyArray[0].WF_Quarter,GlobalMyArray[0].WF_Year,Finaldate,"BulkChange",this.state.Comments,this.props.title).then(function (data)
{

alert('Changed successfully');

window.location.reload();
    
});


this.setState({openbox: true });

//END

}


public closeDialog() {  
  this.setState({ openbox: true })  
}  


public loadItems() {  

   
   let reqApproverIDSID=this.getParam("SID");
   this.setState({myApproverId: reqApproverIDSID});

  }

public async getloginuser()
{

  
  let currentuser= await this._services.getCurrentUser();
  console.log(currentuser);
  let reqApproverID1=await this._services.getItemByEmail(currentuser.Email)
  let reqApproverIDSID1=this.getParam("SID");

  if(reqApproverID1!=reqApproverIDSID1)
  {
   alert('you are not valid user');
   window.location.replace("https://capcoinc.sharepoint.com/sites/SecurityAccessReview/SitePages/NoAcess.aspx");

  }

 
  

}

public getParam(name)
{
 name = name.replace(/[\[]/,"\\\[").replace(/[\]]/,"\\\]");
 var regexS = "[\\?&]"+name+"=([^&#]*)";
 var regex = new RegExp( regexS );
 var results = regex.exec(window.location.href);
 if( results == null )
 return "";
 else
 return results[1];
}

  //Endregion

  private getUserProfileUrl = (loginName: string) => {
    return this._services.getUserProfileUrl(loginName);
  }

  public componentDidUpdate(prevProps: IReactDatatableProps) {
    
    if (prevProps.list !== this.props.list) {
      this.props.onChangeProperty("list");
    }
    else if (this.props.fields != prevProps.fields) {
      this.getSelectedListItems();
    }
  }

  public async getSelectedListItems() {
    let fields = this.props.fields || [];
    let listItems;

    let reqApproverIDSID=this.getParam("SID");
    if (fields.length) {

      //let listItems = await this._services.getListItems(this.props.list, fields);
      if(this.props.AdminView)
      {
        
        listItems = await this._services.getListItems(this.props.list, fields);

        AllListItems=listItems;
      }
      else
      {

        
     listItems=await this._services.getListItemsBasedonApproverID(this.props.list, fields,reqApproverIDSID);
     AllListItems=listItems;
      }
      /** Format list items for data grid */
      listItems = listItems && listItems.map(item => ({
        id: item.Id, ...fields.reduce((ob, f) => {
          ob[f.key] = item[f.key] ? this.formatColumnValue(item[f.key], f.fieldType) : '-';
          return ob;
        }, {})
      }));
      let dataGridColumns: IColumn[] = [...fields].map(f => ({
        key: f.key as string,
        name: f.text,
        fieldName: f.key as string,
        isResizable: true,
        onColumnClick: this.props.sortBy && this.props.sortBy.filter(field => field === f.key).length ? this.handleSorting(f.key as string) : undefined,
        minWidth: 150,
        maxWidth: 400,
        headerClassName: styles.colHeader,
        isMultiline:true
      }));
      this.setState({ listItems: listItems, columns: dataGridColumns });
    }
  }

  private _onConfigure() {
    this.props.context.propertyPane.open();
  }

  
  public formatColumnValue(value: any, type: string) {
    if (!value) {
      return value;
    }
    switch (type) {
      case 'SP.FieldDateTime':
        value = value;
        break;
      case 'SP.FieldMultiChoice':
        value = (value instanceof Array) ? value.join() : value;
        break;
      case 'SP.Taxonomy.TaxonomyField':
        value = value['Label'];
        break;
      case 'SP.FieldLookup':
        value = value['Title'];
        break;
      case 'SP.FieldUser':
        let loginName = value['Name'];
        let userName = value['Title'];
        value = <RenderProfilePicture loginName={loginName} displayName={userName} getUserProfileUrl={() => this.getUserProfileUrl(loginName)}  ></RenderProfilePicture>;
        break;
      case 'SP.FieldMultiLineText':
        //value = <div dangerouslySetInnerHTML={{ __html: 'Tester'}}></div>;
        //value=<div dangerouslySetInnerHTML={{ __html: value.replace(/[\n\r]/g,"<br/>")}}></div>;
        value=<div style={{ whiteSpace: "break-spaces" }}> {value} </div>
        //value=value.replace(/\n/g,"<br>");
        // alert(value);
        // console.log('kalyan');
        // console.log(value);
        break;
      case 'SP.FieldText':
        value = value;
        break;
      case 'SP.FieldComputed':
        value = value;
        break;
      case 'SP.FieldUrl':
        let url = value['Url'];
        let description = value['Description'];
        value = <RenderImageOrLink url={url} description={description}></RenderImageOrLink>;
        break;
      case 'SP.FieldLocation':
        value = JSON.parse(value).DisplayName;
        break;
      default:
        break;
    }
    return value;
  }

  private _getSelectionDetails(): string {
    const selectionCount = this._selection.getSelectedCount();

    switch (selectionCount) {
      case 0:
        return 'No items selected';
      case 1:
        return '1 item selected: ' + (this._selection.getSelection()[0] as any).name;
      default:
        return `${selectionCount} items selected`;
    }
  }

  public formatValueForExportingData(value: any, type?: string) {
    if (!value) {
      return value;
    }
    switch (type) {
      case 'SP.FieldUser':
        let userName = value['Title'];
        value = userName;
        break;
      case 'SP.FieldUrl':
        let url = value['Url'];
        let description = value['Description'];
        value = <a href={url}>{description}</a>;
        break;
      default:
        break;
    }
    return value;
  }

  private exportDataFormatter(fields: Array<IPropertyPaneDropdownOption & { fieldType: string }>, listItems: any[], cellFormatterFn: (value: any, type: string) => any) {
    return listItems && listItems.map(item => ({
      ...fields.reduce((ob, f) => {
        ob[f.text] = item[f.key] ? cellFormatterFn(item[f.key], f.fieldType) : '-';
        return ob;
      }, {})
    }));
  }

  private handlePaginationChange(pageNo: number, rowsPerPage: number) {
    this.setState({ page: pageNo, rowsPerPage: rowsPerPage });
  }

  public handleSearch(event: React.ChangeEvent<HTMLInputElement>) {
    this.setState({ searchText: event.target.value });
  }

  public getTogledata(listItems: any[])
  {

    let listItems1 =listItems.filter(l =>l['ToggleHide']==='Yes');
    return listItems1;
  }

  public filterListItems() {
    let { searchBy, enableSorting,ToggleHide } = this.props;
    let { sortingFields, listItems, searchText } = this.state;
    if (searchText) {
      if (searchBy) {
        listItems = listItems && listItems.length && listItems.filter(l => searchBy.some(field => {
          return (l[field] && l[field].toString().toLowerCase().includes(searchText.toLowerCase()));
        }));
      }
    }
    if (enableSorting && sortingFields) {
      listItems = this.sortListItems(listItems);
    }

    if(ToggleHide)
    {
      listItems = this.getTogledata(listItems);



    }
    return listItems;
  }

  private sortListItems(listItems: any[]) {
    const { sortingFields, sortDirection } = this.state;
    const isAsc = sortDirection === 'asc' ? 1 : -1;
    let sortFieldDetails = this.props.fields.filter(f => f.key === sortingFields)[0];
    let sortFn: (a, b) => number;
    switch (sortFieldDetails.fieldType) {
      case 'SP.FieldDateTime':
        sortFn = (a, b) => ((new Date(a[sortingFields]).getTime() > new Date(b[sortingFields]).getTime()) ? 1 : -1) * isAsc;
        break;
      default:
        sortFn = (a, b) => ((a[sortingFields] > b[sortingFields]) ? 1 : -1) * isAsc;
        break;
    }
    listItems.sort(sortFn);
    return listItems;
  }

  private paginateFn = (filterItem: any[]) => {
    let { rowsPerPage, page } = this.state;
    return (rowsPerPage > 0
      ? filterItem.slice((page - 1) * rowsPerPage, (page - 1) * rowsPerPage + rowsPerPage)
      : filterItem
    );
  }

  private handleSorting = (property: string) => (event: React.MouseEvent<unknown>, column: IColumn) => {
    property = column.key;
    let { sortingFields, sortDirection } = this.state;
    const isAsc = sortingFields && sortingFields === property && sortDirection === 'asc';
    let updateColumns = this.state.columns.map(c => {
      return c.key === property ? { ...c, isSorted: true, isSortedDescending: (isAsc ? false : true) } : { ...c, isSorted: false, isSortedDescending: true };
    });
    this.setState({ sortDirection: (isAsc ? 'desc' : 'asc'), sortingFields: property, columns: updateColumns });
  }

  private _onRenderRow: IDetailsListProps['onRenderRow'] = props => {
    const customStyles: Partial<IDetailsRowStyles> = {};
    if (props) {
      if (props.itemIndex % 2 === 0) {
        customStyles.root = { backgroundColor: this.props.evenRowColor };
      }
      else {
        customStyles.root = { backgroundColor: this.props.oddRowColor };
      }
      return <DetailsRow {...props} styles={customStyles} />;
    }
    return null;
  }

  private changeClientName(data: any): void {

    this.setState({ Comments: data.target.value });

  }

  public removewithfilter(arr) {
    let outputArray = arr.filter(function(v, i, self)
    {
         
        // It returns the index of the first
        // instance of each value
        return i == self.indexOf(v);
    });
     
    return outputArray;
}



  
  public render(): React.ReactElement<IReactDatatableProps> {
    let filteredItems = this.filterListItems();
    let { list, fields, enableDownloadAsCsv, enableDownloadAsPdf, enablePagination, displayMode, enableSearching, title, evenRowColor, oddRowColor } = this.props;
    let { columns } = this.state;
    console.log(columns);
    console.log('kalyan');
    console.log(fields);
    console.log(this._selection.getSelection());

    
    
    let filteredPageItems = enablePagination ? this.paginateFn(filteredItems) : filteredItems;

    return (
      <div className={styles.reactDatatable}>
        {
          this.props.list == "" || this.props.list == undefined ?
            <Placeholder
              iconName='Edit'
              iconText='Configure your web part'
              description={strings.ConfigureWebpartDescription}
              buttonLabel={strings.ConfigureWebpartButtonLabel}
              hideButton={displayMode === DisplayMode.Read}
              onConfigure={this._onConfigure} /> : <><>
                <WebPartTitle
                  title={title}
                  displayMode={DisplayMode.Read}
                  updateProperty={() => { }}>
                </WebPartTitle>
                {list && fields && fields.length ?
                  <div>
                    <Grid container className={styles.dataTableUtilities}>
                      <Grid item xs={6} className={styles.downloadButtons}>
                        {enableDownloadAsCsv
                          ? <ExportListItemsToCSV
                            columnHeader={columns.map(c => c.name)}
                            listName={list}
                            description={title}
                            dataSource={() => this.exportDataFormatter(fields, filteredItems, csvCellFormatter)}
                          /> : <></>}
                        {enableDownloadAsPdf
                          ? <ExportListItemsToPDF
                            listName={list}
                            title={title}
                            columns={columns.map(c => c.name)}
                            oddRowColor={oddRowColor}
                            evenRowColor={evenRowColor}
                            dataSource={() => this.exportDataFormatter(fields, filteredItems, pdfCellFormatter)} />
                          : <></>}
                      </Grid>
                      <Grid container justify='flex-end' xs={6}>
                        {enableSearching ?
                          <TextField 
                            onChange={this.handleSearch.bind(this)}
                            placeholder="Search"
                            className={styles.txtSearchBox} />
                          : <></>}
                      </Grid>
                    </Grid>
                    <div id="generateTable">
                      <DetailsList 
                        items={filteredPageItems}
                        columns={columns}
                        selectionMode={this.props.AdminView?SelectionMode.none:SelectionMode.multiple}
                        layoutMode={DetailsListLayoutMode.justified}
                        isHeaderVisible={true}
                        onRenderRow={this._onRenderRow}
                        selection={this._selection}
            selectionPreservedOnEmptyClick={true}
           ariaLabelForSelectionColumn="Toggle selection"
            ariaLabelForSelectAllCheckbox="Toggle selection for all items"
            checkButtonAriaLabel="select row"
                      />
                      <div>
                        {this.props.enablePagination ?
                          <Pagination
                            currentPage={this.state.page}
                            totalItems={filteredItems.length}
                            onChange={this.handlePaginationChange.bind(this)}
                          />
                          : <></>}
                      </div>
                    </div>


                  </div> : <MessageBar>
                    {strings.ListFieldValidation}
                  </MessageBar>}</>
            </>
        }

        <br></br>
        <br></br>
        <br></br>
        <br></br>
        <br></br>
        <br></br>

<table hidden={this.props.AdminView}>
  <tr>
    <td>
    <PrimaryButton type='button' onClick={this.Approve.bind(this)}>Approve</PrimaryButton>
   
    </td>
    <td>
    <PrimaryButton name='TestButton' onClick={this.Remove.bind(this)}>Remove</PrimaryButton>
     </td>
    <td>
    <PrimaryButton name='TestButton' onClick={this.changerecord.bind(this)} >Change Approver</PrimaryButton>
    </td>
  </tr>
  
</table>

<Dialog hidden={this.state.openbox}>

<Stack horizontal tokens={sectionStackTokens} className={styles.myStyles}>
<StackItem className={styles.labelstyles}>
<label>Employee ID</label>
</StackItem>
<StackItem>
<label>{this.state.EmpID}</label><br></br>
</StackItem>
</Stack>
<Stack horizontal tokens={sectionStackTokens} className={styles.myStylesSpaces}>
<StackItem className={styles.labelstyles}>
<label>Employee Name</label>
 </StackItem>
<StackItem>
<label>{this.state.EmpName}</label><br></br>
</StackItem>
</Stack>
<Stack horizontal tokens={sectionStackTokens} className={styles.myStylesSpaces}>
<StackItem className={styles.labelstyles}>
<label>Comments *</label><br></br>
 </StackItem>
<StackItem>
<TextField multiline={true} value={this.state.Comments} onChange={this.changeClientName.bind(this)} ></TextField>
</StackItem>
</Stack>
<br></br>
<Stack horizontal tokens={sectionStackTokens} className={styles.myStylesSpaces}>
<StackItem>
<PrimaryButton name='TestButton' onClick={this.savechangetest.bind(this)}>Save</PrimaryButton>
  </StackItem>
  <StackItem>
  <PrimaryButton name='TestButton' onClick={this.closeDialog.bind(this)}>Cancel</PrimaryButton>
  </StackItem>
</Stack>

</Dialog>
</div >
      
    );

    
  }

 
}

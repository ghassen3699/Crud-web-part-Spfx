import * as React from 'react';
import styles from './FreshList.module.scss';
import { IFreshListProps } from './IFreshListProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { Toolbar } from '@pnp/spfx-controls-react/lib/Toolbar';
import { DefaultButton, PrimaryButton } from '@fluentui/react/lib/Button';
import { Panel } from '@fluentui/react/lib/Panel';
import { useBoolean } from '@fluentui/react-hooks';
import { ListView, IViewField, SelectionMode, GroupOrder, IGrouping } from "@pnp/spfx-controls-react/lib/ListView";
import * as _ from 'lodash';
import { DisplayMode } from '@microsoft/sp-core-library';
import { DynamicForm } from "@pnp/spfx-controls-react/lib/DynamicForm";
import { WebPartTitle } from "@pnp/spfx-controls-react/lib/WebPartTitle";
import { Placeholder } from "@pnp/spfx-controls-react/lib/Placeholder";
import { Pagination } from "@pnp/spfx-controls-react/lib/pagination";







export default class FreshList extends React.Component<IFreshListProps, {}> {

  public state = {
    data:[],  // data of list
    openAddPanel: false,   // open and close the property pane of add item
    openUpdatePanel: false,  // open and close the property pane of update item
    listID: '',   // the ID of the list 
    itemID: 0  // item selected id 
  };

  // variables of webpart title
  public displayMode: DisplayMode;
  public updateProperty: (value: string) => void;



  // https://x2r2q.sharepoint.com/sites/FreshComponent/_api/Web/Lists/getbytitle('Employees')/items
  // https://projetaziz.sharepoint.com/sites/Alight_ghassenProject_freshWebPart/_api/Web/Lists/getbytitle('Employees')/items?$select=ID,Title


  // get all items from list with API properties
  private getAllItemsAPI = async(ApiURL, ColumnSelectedProperties, numberOfElement) => {
    var listData = [];

    // Get all items from list with fetch 
    var result = await fetch(ApiURL, {
        method: 'GET',
        headers: {
              "Accept": "application/json;odata=verbose",
              "Content-type": "application/json;odata=verbose",
            },
      })
      .then((response) => { 
        return response.json().then((data) => {
            return data.d.results;
        }).catch((err) => {
            console.log(err);
        });
    });
    if (result !== undefined)
    {
      // get all colummn selected
      const selectedColumn = this.getSelectedColumn(ColumnSelectedProperties);
      
      // Get data from column selected
      result.map(item => {
        let secondList = []
        for(let i = 0; i < selectedColumn.length; i++){
          if (item[selectedColumn[i]] !== undefined){
            secondList.push([selectedColumn[i],item[selectedColumn[i]]]);
          }
        }

        // convert the list to an object and push it to the listData
        var objectData = _.fromPairs(secondList);
        listData.push(objectData);
        
      })

      // limit the data with the number of element selected in property pane
      if (numberOfElement < listData.length){
        listData = listData.slice(0,numberOfElement)  ;    
      }
      this.setState({data:listData});
    }

  };


  

  // Delete item from list 
  private async deleteItem (Id, listName){
    if (Id !== 0){
      let listeData = sp.web.lists.getByTitle(listName);
      let response = await listeData.items.getById(Id).delete();
      location.reload()
    }else {
      alert('Sélectionner une ligne pour le supprimer')
    }
  };




  // convert input from string to an array of column selected
  private getSelectedColumn = (columnSelected) => {
    const myColumns = columnSelected.split(",");
    return myColumns;
  };



  // get the name of list and convert that to an ID
  private getListID = async (listNamePropertie) => {
    if (listNamePropertie !== ''){
      let list = await sp.web.lists.getByTitle(listNamePropertie).get();
      let ID = list.Id;
      this.setState({listID:ID});
    }
  };



  // Selection methode in table
  private _getSelection = (items:any[]) => {
    if (items.length !== 0){
      this.setState({itemID:items[0]['ID']})
    }else {
      this.setState({itemID:0})
    }
  };


  // private _getPage(page: number){
  //   console.log('Page:', page);
  // }



  // Action buttons of Toolbar before select Item in list
  public BeforeSelected = {
    'group1': {
      'action1': {
        title: 'Ajouter',
        iconName: 'Add',
        onClick: () => this.openAddPanel()
      },
      'action2': {
        title: 'Modifier',
        iconName: 'Edit',
        onClick: () => this.openUpdatePanel()
      },
      'action3': {
        title: 'Supprimer',
        iconName: 'Delete',
        onClick: () => { this.deleteItem(this.state.itemID, this.props.listName) }
      }
    }
  };
  

  // open and close Add panel
  private openAddPanel(){
    const openAddPanel = this.state.openAddPanel;
    this.setState({openAddPanel:!openAddPanel});   
    if (this.state.openAddPanel){
      location.reload()
    }
  };


  // open and close update panel
  private openUpdatePanel(){
    if (this.state.itemID !== 0){
      const openUpdatePanel = this.state.openUpdatePanel;
      this.setState({openUpdatePanel:!openUpdatePanel});
    }else {
      alert('Sélectionner une ligne pour le Modifier')
    }
    if (this.state.openUpdatePanel){
      location.reload()
    }
  };

  // open webpart property pane if one of data doesn't exist
  private _onConfigure = () => {
    // Context of the web part
    this.props.context.propertyPane.open();
  };

  // initialise the state
  componentDidMount(): void {
    // get all data from list and setstate them
    var listUrlAPI = this.props.context.pageContext.web.absoluteUrl + this.props.listUrlAPI;
    this.getAllItemsAPI(listUrlAPI, this.props.columnSelected, this.props.numberOfElement);
    // get the ID of the list
    this.getListID(this.props.listName);
  };


  // the name of Add item property pane 
  public propertiePaneAddName = "Add item to " + this.props.listName; 
  // the name of update item property pane
  public propertyPaneUpdateName = "Update item to " + this.props.listName;

  // // update state
  // componentDidUpdate(prevProps, prevState, snapshot?: any): void {
  //   if (prevState.data !== this.state.data){
  //     var listUrlAPI = this.props.context.pageContext.web.absoluteUrl + prevProps.listUrlAPI;
  //     this.getAllItemsAPI(listUrlAPI,prevProps.columnSelected,prevProps.numberOfElement)
  //   } 
  // };


  // Render methode 
  public render(): React.ReactElement<IFreshListProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName,
      listUrlAPI,
      numberOfElement,
      themeColor,
      webPartName,
      columnSelected,
      listName,
      context,
    } = this.props;

    if ((listUrlAPI === undefined) || (columnSelected === undefined) || (listName === undefined) || (numberOfElement === undefined)){
      return(
        <Placeholder iconName='Edit'
          iconText='Configurer votre web part'
          description="S'il vous plais configurer votre web part."
          buttonLabel='Configurer'
          onConfigure={this._onConfigure}
        />
      );
    }else {
      return (
        <section className={`${styles.freshList} ${hasTeamsContext ? styles.teams : ''}`}>
  
          {console.log(this.state.itemID)}
          {/* toolbar of web component  */}
            <WebPartTitle 
              displayMode={this.displayMode}
              title={webPartName}
              updateProperty={this.updateProperty} 
            />
          {/* ************************* */}
  

          {/* toolbar of web component  */}
            <Toolbar actionGroups={this.BeforeSelected}/>
          {/* ************************* */}
  
  
          {/* First panel for add new item to the list */}
            <ListView
              items={this.state.data}
              compact={false}
              selectionMode={SelectionMode.single}
              selection={this._getSelection}
              stickyHeader={true}

            />
          {/* ***************************************** */}
          
  
          {/* First panel for add new item to the list */}
            <Panel
              headerText={this.propertiePaneAddName}
              isOpen={this.state.openAddPanel}
              closeButtonAriaLabel="Close"
              isFooterAtBottom={true}
              onDismiss={() => this.openAddPanel()}>
              <h3>New item.</h3>
  
              {/* add new item form */}
              <DynamicForm
                key={this.state.listID} 
                context={this.props.context} 
                listId={this.state.listID}>
              </DynamicForm>
              {/* ****************** */}
  
            </Panel>
          {/* ****************************************** */}
  
  
          {/* Second panel for update item in the list */}
            <Panel
              headerText={this.propertiePaneAddName}
              isOpen={this.state.openUpdatePanel}
              closeButtonAriaLabel="Close"
              isFooterAtBottom={true}
              onDismiss={() => this.openUpdatePanel()}>
              <h3>Update item.</h3>
  
              {/* update item form */}
              <DynamicForm
                key={this.state.listID} 
                context={this.props.context} 
                listId={this.state.listID}
                listItemId={this.state.itemID}  >
              </DynamicForm>
              {/* ****************** */}
  
            </Panel>
          {/* ****************************************** */}

          {/* <Pagination
            currentPage={3}
            totalPages={13} 
            onChange={(page) => this._getPage(page)}
            limiter={3} // Optional - default value 3
            hideFirstPageJump // Optional
            hideLastPageJump // Optional
            limiterIcon={"Emoji12"} // Optional
          /> */}
          
  
        </section>
      );
    };
  };
}

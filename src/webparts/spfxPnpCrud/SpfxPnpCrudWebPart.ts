import { Version } from "@microsoft/sp-core-library";
import "bootstrap/dist/css/bootstrap.min.css";
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
// import type { IReadonlyTheme } from '@microsoft/sp-component-base';
// import { escape } from '@microsoft/sp-lodash-subset';

// import styles from './SpfxPnpCrudWebPart.module.scss';
import * as strings from "SpfxPnpCrudWebPartStrings";
import * as pnp from "sp-pnp-js";

export interface ISpfxPnpCrudWebPartProps {
  description: string;
  
}

export default class SpfxPnpCrudWebPart extends BaseClientSideWebPart<ISpfxPnpCrudWebPartProps> {
  public render(): void {
    this.domElement.innerHTML = `
    <div>
    <div class="container bg-body-secondary">
      <table border='5'>
        <tr>
          <td>Student Id</td>
          <td><input type='text' id='studentId'/></td>
          <td><input type='submit' id='btnRead' value='GetDetails'/></td>
        </tr>
  
        <tr>
          <td>Student Name</td>
          <td><input type='text' id='txtstudentName'/></td>
        </tr>
  
        <tr>
          <td>Student department</td>
          <td><input type='text' id='txtstudentDept'/></td>
        </tr>
  
        <tr>
          <td>Student City</td>
          <td><input type='text' id='txtstudentCity'/></td>
        </tr>
  
        <tr>
          <td>
            <input type='submit' value='Insert' id='btnInsert'/>
            <input type='submit' value='Update' id='btnUpdate'/>
            <input type='submit' value='Delete' id='btnDelete'/>
            <input type='submit' value='show All Records' id='btnReadAll'/>
          </td>
        </tr>
      </table>
    </div>
    <div id='MsgStatus'></div>
    <h2>Get All List Items</h2>
    <hr/>
    <div id="spListData" />
  </div>
  `;
  this.bindEvent();
  this.readAllItems();
  }

  public readAllItems():void{
    let html: string = '<table border=1 width=100% style="bordercollapse: collapse;">';
    html += `<th>Title</th><th>Name</th><th>City</th><th>Depaertment</th>`;
    pnp.sp.web.lists.getByTitle('Students').items.get().then((items:any[])=>{
        items.forEach(function(item){
          html+= `
          <tr>
          <td>${item["Title"]}</td>
          <td>${item["Name"]}</td>
          <td>${item["City"]}</td>
          <td>${item["Depaertment"]}</td>
          </tr>
          `;
        })
        html+=`</table>`;
        const allItems: Element | null = this.domElement.querySelector('#spListData');
        allItems !== null ? allItems.innerHTML = html  : "";
    })

    
  }

  private bindEvent() :void{
    this.domElement.querySelector('#btnInsert')?.addEventListener('click' ,()=>{
      this.insertStudent();
    })
    this.domElement.querySelector('#btnRead')?.addEventListener('click' ,()=>{
      this.readStudent();
    })
    this.domElement.querySelector('#btnUpdate')?.addEventListener('click' ,()=>{
      this.updateStudent();
    })
    this.domElement.querySelector('#btnDelete')?.addEventListener('click' ,()=>{
      this.deleteStudent();
    })

  }


  // delete 

  private  deleteStudent():void{
    let studentID  = parseInt((document.getElementById("studentId") as HTMLInputElement)?.value);
    pnp.sp.web.lists.getByTitle('Students').items.getById(studentID).delete();
    alert("deleted sucefully");
    

  }
 // update the data in  the tables
    private updateStudent(): void{
      var StudentId = (document.getElementById('studentId') as HTMLInputElement)?.value;
      var StudentName = (document.getElementById('txtstudentName') as HTMLInputElement)?.value;
      var StudentDept = (document.getElementById('txtstudentDept') as HTMLInputElement)?.value;
      var StudentCity = (document.getElementById('txtstudentCity') as HTMLInputElement)?.value;
       console.log(StudentId);
       console.log(StudentName);
       console.log(StudentDept);
       console.log(StudentCity);
    
       //const siteurl: string = this.context.pageContext.site.absoluteUrl + "/_api/web/lists/getbytitle('Students')/items";
    
       let studentID  = parseInt((document.getElementById("studentId") as HTMLInputElement)?.value);
        pnp.sp.web.lists.getByTitle('Students').items.  getById(studentID).update({
          Title: StudentId,
          Name: StudentName,
          Depaertment: StudentDept ,
          City: StudentCity
        }).then(response =>{
          console.log( "details added ",response)
    
        }).catch(err=>console.log("error occured" ,err))
        ;
      
    }


  private readStudent(): void {
    // CODE WORKING 
    // PUT INPUT AS 1 - 9 
    let StudentId  = parseInt((document.getElementById("studentId") as HTMLInputElement)?.value);
     console.log("here is student id",StudentId)
    
    pnp.sp.web.lists.getByTitle("Students").items.getById(StudentId).get().then((item) => {
      (document.getElementById('txtstudentName') as HTMLInputElement).value = item.Name;
      (document.getElementById('txtstudentCity') as HTMLInputElement).value = item.City;
      (document.getElementById('txtstudentDept') as HTMLInputElement).value  = item.Depaertment;
    })
   
  }
  
  private insertStudent() :void{
  var StudentId = (document.getElementById('studentId') as HTMLInputElement)?.value;
  var StudentName = (document.getElementById('txtstudentName') as HTMLInputElement)?.value;
  var StudentDept = (document.getElementById('txtstudentDept') as HTMLInputElement)?.value;
  var StudentCity = (document.getElementById('txtstudentCity') as HTMLInputElement)?.value;
   console.log(StudentId);
   console.log(StudentName);
   console.log(StudentDept);
   console.log(StudentCity);

   //const siteurl: string = this.context.pageContext.site.absoluteUrl + "/_api/web/lists/getbytitle('Students')/items";


    pnp.sp.web.lists.getByTitle('Students').items.add({
      Title: StudentId,
      Name: StudentName,
      Depaertment: StudentDept ,
      City: StudentCity
    }).then(response =>{
      console.log( "item added ",response)

    }).catch(err=>console.log("error occured" ,err))
    ;
  }
  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription,
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField("description", {
                  label: strings.DescriptionFieldLabel,
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}

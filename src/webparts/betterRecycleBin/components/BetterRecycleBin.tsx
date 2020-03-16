import * as React from 'react';
import styles from './BetterRecycleBin.module.scss';
import { IBetterRecycleBinProps } from './IBetterRecycleBinProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { sp, Web, Site, SiteGroup, SiteGroups } from "@pnp/sp/presets/all";
import { SiteUser } from '@pnp/sp/site-users';
import ReactTable from 'react-table-6';
import 'react-table-6/react-table.css';
import { FontAwesomeIcon } from '@fortawesome/react-fontawesome';
import { faSquare, faCheckSquare } from '@fortawesome/free-regular-svg-icons';
import MomentReact  from 'react-moment';
import {Moment} from 'moment';
import fetchStream from 'fetch-readablestream';
import axios from 'axios';
import * as moment from 'moment';

export class TrashItem {
  public Title:string;
  public Deleted:Date;
  public DeletedBy:string;
  public Path:string;
  public Selected:boolean;
  public Id: string;

  constructor(item:any) {
      if (item.Title) {
          this.Title = item.Title;
      }
      else if (item.LeafName) {
          this.Title = item.LeafName;            
      }
      else {
          this.Title = "";
      }
      this.Deleted = item.DeletedDate;
      this.DeletedBy = item.DeletedByName;
      this.Path = item.DirName;
      this.Selected = false;
      this.Id = item.Id;
  }
}

export default class BetterRecycleBin extends React.Component<IBetterRecycleBinProps, {siteName:string; log: string[]; trash: any[]; siteUrl: string; showTool: boolean}> {
  constructor(props) {
    super(props);
    sp.setup({
      spfxContext: this.context
    });
    this.state={
      siteName: "",
      log: [],
      trash: [],
      siteUrl: "",
      showTool: true
    };
    this.logMessage = this.logMessage.bind(this);
    this.getRecycleBin = this.getRecycleBin.bind(this);
    this.selectItem = this.selectItem.bind(this);
    this.restoreSelected = this.restoreSelected.bind(this);
    this.refresh = this.refresh.bind(this);
  }

  private logMessage(message:string) {
    this.setState({log: [...this.state.log, message]});
  }

  private async getDigest(){
    return await axios.post(this.state.siteUrl + "/_api/contextinfo", {
        headers: { 'accept': 'application/json;odata=nometadata' },
    }).then(response => {
        return response.data.FormDigestValue;
    });
  }

  private refresh() {
    this.getRecycleBin();
  }

  private async restoreSelected() {
    this.setState({log: []});

    let newTrash = [];
    for (let x = 0; x < this.state.trash.length; x++) {
      let item = this.state.trash[x];
      if (item.Selected) {
        let digest = await this.getDigest();
        let response = await fetch(this.state.siteUrl + "/_api/site/RecycleBin('" + item.Id + "')/restore()", {headers:{"accept": "application/json", 'X-RequestDigest': digest}, method:"POST"});
        let chunks = await this.readAllChunks(response.body);
        let stringValue = "";
        chunks.forEach(chunk => {
          stringValue += new TextDecoder("utf-8").decode(chunk);
        });
        let message = JSON.parse(stringValue);
        let error = message['odata.error'];
        if (error) {
            this.logMessage(error.message.value); 
            newTrash.push(item);           
        }
        else {
            this.logMessage("Successfully restored " + item.Title);
        }
      }
      else {
        newTrash.push(item);
      }
    }

    this.setState({trash: newTrash});
  }

  
  private selectItem(props) {
    let newTrash = [];

    this.state.trash.forEach(t=>{
        if  (t.Id === props.original.Id) {
            if (t.Selected) {
                t.Selected = false;
            }
            else {
                t.Selected = true;
            }                
        }
        newTrash.push(t);
    });

    this.setState({trash: newTrash});
}

  private async getRecycleBin() {
    this.setState({trash: []});
    fetchStream(this.state.siteUrl + "/_api/site/RecycleBin", {headers:{"accept": "application/json"}})
    .then(response => this.readAllChunks(response.body))
    .then(chunks => {
      let stringValue = "";
      chunks.forEach(chunk => {
        stringValue += new TextDecoder("utf-8").decode(chunk);
      });
      let recycleBin = JSON.parse(stringValue);
      this.logMessage("Found " + recycleBin.value.length + " items in the Recycle Bin");
      recycleBin.value.forEach(t=>{
        let newItem = new TrashItem(t);
        this.setState({trash: [...this.state.trash, newItem]});
    });
    });
  }

  private readAllChunks(readableStream) {
    const reader = readableStream.getReader();
    const chunks = [];
   
    function pump() {
      return reader.read().then(({ value, done }) => {
        if (done) {
          return chunks;
        }
        chunks.push(value);
        return pump();
      });
    }
   
    return pump();
  }

  public async componentDidMount() {    
    const site = await sp.site.get();
    const rootWeb = await Web(site.Url).get();
    let currentUser = await sp.web.currentUser.get();
    let showTool = currentUser.IsSiteAdmin;
    let siteUrl = rootWeb.ServerRelativeUrl;
    this.setState({siteUrl: siteUrl});
    this.setState({showTool: showTool});
    if (showTool) {
      this.logMessage("Getting environment");
      this.setState({siteName: rootWeb.Title});
      this.logMessage("Getting Recycle Bin");
      this.getRecycleBin();
    }
  }
    
  public render(): React.ReactElement<IBetterRecycleBinProps> {
    const columns = [
      {
          Header: '',
          accessor: 'Selected',
          maxWidth: 50,
          filterable: false,                
          Cell: props=> <span className={styles.recycleBin}>{props.value &&
              <span onClick={()=>this.selectItem(props)}><FontAwesomeIcon icon={faCheckSquare} /></span> ||
              <span onClick={()=>this.selectItem(props)}><FontAwesomeIcon icon={faSquare} /></span>}</span>                
      },
      {
          Header: 'Title',
          accessor: 'Title', 
          filterMethod: (filter, row) => {
              let rowValue = row[filter.id].toString().toLowerCase();
              let filterValue = filter.value.toString().toLowerCase();
              if (rowValue.indexOf(filterValue) >= 0) {
                  return true;                        
              }
              else {
                  return false;
              }
          }
      },
      {
          Header: 'Path',
          accessor: 'Path',
          filterMethod: (filter, row) => {
              let rowValue = row[filter.id].toString().toLowerCase();
              let filterValue = filter.value.toString().toLowerCase();
              if (rowValue.indexOf(filterValue) >= 0) {
                  return true;                        
              }
              else {
                  return false;
              }
          }
      },
      {
          Header: 'Deleted',
          accessor: 'Deleted',
          maxWidth:200,
          Cell: props=> <MomentReact format={"h:mm:ssa MM/DD/YYYY"}>{props.original.Deleted}</MomentReact>,
          filterMethod: (filter, row) => {
              let filterDate = moment(row[filter.id]).format("h:mm:ssa MM/DD/YYYY");
              if (filterDate.indexOf(filter.value) >= 0) {
                  return true;
              }
              else {
                  return false;
              }
          }
      },
      {
          Header: 'Deleted By',
          accessor: 'DeletedBy',
          maxWidth: 200,
          filterMethod: (filter, row) => {
              let rowValue = row[filter.id].toString().toLowerCase();
              let filterValue = filter.value.toString().toLowerCase();
              if (rowValue.indexOf(filterValue) >= 0) {
                  return true;                        
              }
              else {
                  return false;
              }
          }
      }
    ];

    return (
      <div className={ styles.betterRecycleBin }>
        {this.state.showTool &&
        <React.Fragment>
          <div className={styles.description}>
            <h2>Better Recycle Bin - Environment: {this.state.siteName} {this.state.siteUrl}</h2>
            <h3>{this.props.description}</h3>
          </div>
          {this.state.trash.length > 0 &&
            <React.Fragment>
              <div className={styles.table}>
                  <ReactTable 
                      data={this.state.trash}
                      columns={columns}
                      defaultPageSize={10}
                      filterable={true}
                  />
              </div>
              <div className={styles.buttons}>
                  <button type="button" onClick={this.restoreSelected} className={styles.button}>Restore Selected</button>
                  <button type="button" onClick={this.refresh}>Refresh List</button>
              </div>
              <div className={styles.log}>
              {this.state.log.length > 0 &&
                  <ul>
                      {this.state.log.map(l=>(
                          <li>{l}</li>
                      ))}
                  </ul>
              }
              </div>
            </React.Fragment>
          }
        </React.Fragment>
        ||
        <React.Fragment>
          <h1 className={styles.description}>Sorry, this tool is for Site Collection Admins only!</h1>
        </React.Fragment>
        }
      </div>
    );
  }
}

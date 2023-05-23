import * as React from 'react';
import styles from './OverviewOfApplications.module.scss';
import { IOverviewOfApplicationsProps } from './IOverviewOfApplicationsProps';
import { sp, ISPConfiguration } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-users/web"
import { Toggle } from 'office-ui-fabric-react';

export default class OverviewOfApplications extends React.Component<IOverviewOfApplicationsProps, {}> {

  showInfo : boolean = false;
  loaded : boolean = false;
  listName : string = "Přehled aplikací intranetu FN Brno";
  listItems : any[] = [];
  adminID : any;
  admin : {};
  
  private async fetchSP() : Promise<any> {
    const config: ISPConfiguration = {
      sp: {
        baseUrl: "https://fnembrno.sharepoint.com/sites/PrehledAplikaci/",
      },
    };

    sp.setup(config);

    return await sp.web.lists.getByTitle(this.listName).items();
    
  }

  private async getUser(id : any) : Promise<any> {
    return await sp.web.siteUsers.getById(id).get();
  }

  // TODO je zadouci aby se tato data fetchovala pri mountu? a kdy vzdycky se provede mount?
  public componentDidMount() : void {

    console.log("MOUNTED");

    this.fetchSP().then((response) => {
      this.listItems = response;
      this.loaded = true;
      this.forceUpdate();

      for (const item of this.listItems) {

        if (item.SpravceId != null) {
          this.getUser(item.SpravceId).then((user) => {
            this.admin = {name : user.Title, email : user.Email};

            if (item.SpravceId == null) {
              this.admin = null;
            }

            item.SpravceId = this.admin;

            if (this.listItems.length -1 == this.listItems.indexOf(item)) {
              this.forceUpdate();
            }

          }).catch((error) => {console.log(error)});
        }
      }
    }).catch((error) => {console.log(error)});
  }

  public render(): React.ReactElement<IOverviewOfApplicationsProps> {
    const {} = this.props;
    console.log("RENDERED");

    return (
      <div className={styles.overviewOfApplications}>

        <div className={styles.layout}>

          <Toggle onText="Skrýt správce databáze" offText="Zobrazit správce databáze" onChange={() => {this.showInfo = !this.showInfo; this.forceUpdate()}} defaultChecked={false}></Toggle>
          
          {this.loaded ? this.listItems.map((item) => 
            <div onClick={() => {window.open(item.Odkaz.Url, "_blank")}} className={styles.apps}>

              <img src={JSON.parse(item.Ikona).serverUrl + JSON.parse(item.Ikona).serverRelativeUrl} width="25px" height="25px"/>

              <div>
              {item.NazevDatabaze}
              <br/>
              {this.showInfo ? <span className={styles.info}><b>Správce databáze</b><br/>{item.SpravceId ? item.SpravceId.name : null}<br/>{item.SpravceId ? item.SpravceId.email : null}<br/><b>Telefon: {item.Telefon ? item.Telefon : null}</b></span> : null}
              </div>

            </div>) : "Načítání ..."}
        </div>
      </div>
    );
  }
}

import * as React from 'react';
import styles from './LrWebpartHideShow.module.scss';
import { ILrWebpartHideShowProps } from './ILrWebpartHideShowProps';
import '@pnp/sp/webs';
import '@pnp/sp/profiles';
import { sp } from '@pnp/sp';
import '@pnp/sp/folders';
import '@pnp/sp/files';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import { ImageFit } from 'office-ui-fabric-react';
import { Carousel, CarouselButtonsDisplay, CarouselButtonsLocation, CarouselIndicatorShape, CarouselIndicatorsDisplay } from "@pnp/spfx-controls-react/lib/Carousel";


export interface ILrWebpartHideShowState{ 
  currentUser: string;
  selectedData:any;
  currentUserOfficeLocation: string;
  currentUserCountry: string;
  selectedSitePage:any;
  sitePagesCount:boolean;
  message: string;
  isLoading: boolean;
}

export default class LrWebpartHideShow extends React.Component<ILrWebpartHideShowProps, ILrWebpartHideShowState> {

  constructor(props: ILrWebpartHideShowProps){
    super(props);
    this.state = {
      currentUser: null,
      selectedData:[],
      currentUserOfficeLocation: null,
      currentUserCountry: null,
      selectedSitePage:[],
      sitePagesCount:false,
      message: 'There are no items available based on your office location or country',
      isLoading:true,
    };
  } 

  public componentDidMount(){ 
 
    this.getUserLocation().then((location) =>{
      // console.log(location);
      this.setState({
        currentUserOfficeLocation: location
      });
      console.log(this.state.currentUserOfficeLocation);
      this.getUserCountry().then((country) => {
        this.setState({
          currentUserCountry : country
        }); 
        console.log(this.state.currentUserCountry);  
        
        this.getSitePage().then(pages => {
          //console.log(pages);
          this.setState({
            selectedSitePage: pages,
            isLoading:false,
          });
          console.log(this.state.selectedSitePage);
          console.log(this.state.isLoading);
        });
      });      
    });        
  }

  public getSitePage(): Promise<string[]> {
    let sitePageData: any = [];   
    return sp.web.lists.getByTitle("Site Pages").items.select("Title","Description", "FileRef", "OfficeLocation", "Country", "ImageURL").getAll().then((data) => {
      data.forEach((index) => {
        if(index.OfficeLocation !== null){
          if(index.OfficeLocation.toLowerCase() == this.state.currentUserOfficeLocation.toLowerCase()){
            sitePageData = [...sitePageData,index];
          }
        }
        else if(index.Country !== null){
          if(index.Country.toLowerCase() == this.state.currentUserCountry.toLowerCase()){
            sitePageData = [...sitePageData,index];
          }
        }                 
      });    
      return sitePageData;   
    });      
  }

  public getUserCountry(): Promise<string>{      
    let userCountry = null;
    return sp.profiles.myProperties.get().then(result => {
      // console.log("result - " + result);
      let properties = result.UserProfileProperties;     
      for(var i = 0; i < properties.length; i++){        
        if(properties[i]["Key"] == "Location"){
          userCountry = properties[i]["Value"];
        }
      }
      return userCountry;      
    });  
  }

  public getUserLocation(): Promise<string>{ 
    let userOfficeLocation = null;  
    
    return sp.profiles.myProperties.get().then(result => {
      // console.log("result - " + result);
      let properties = result.UserProfileProperties;     
      for(var i = 0; i < properties.length; i++){
        if(properties[i]["Key"] == "SPS-Location"){
          userOfficeLocation = properties[i]["Value"];
        }
        
      }
      return userOfficeLocation;      
    });  
  }

  public render(): React.ReactElement<ILrWebpartHideShowProps> {   
    return(
      <div className= {styles.container}>       
        <div className={styles.row}>
          {this.state.isLoading === false ? 
          <div className={styles.column}>
            {this.state.selectedSitePage.length > 0  ? 
              <Carousel 
                buttonsLocation={CarouselButtonsLocation.center}
                buttonsDisplay={CarouselButtonsDisplay.buttonsOnly}
                indicatorShape={CarouselIndicatorShape.circle}
                pauseOnHover={true}
                isInfinite={true}
                interval={2000}
                element={this.state.selectedSitePage.map((page,i) => ({
                  imageSrc:page.ImageURL.Url,
                  title: page.Title,
                  description: page.Description,
                  url: page.FileRef,
                  target: "_self",
                  imageFit: ImageFit.cover
                }))}                               
              /> : <span className='span-font'>{this.state.message}</span> 
            }         
          </div> : <span className='span-font'>Loading...</span>} 
        </div>
      </div>
    );    
  }
}

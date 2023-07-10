import * as React from 'react';
import { BaseWebComponent } from '@pnp/modern-search-extensibility';
import * as ReactDOM from 'react-dom';
import { PageContext } from '@microsoft/sp-page-context';
import {MSGraphClientFactory, SPHttpClient} from "@microsoft/sp-http";
import { initializeFileTypeIcons, getFileTypeIconProps } from '@uifabric/file-type-icons';
import { Icon, CommandBarButton, IconButton, TooltipHost } from 'office-ui-fabric-react';
import { followDocument, unFollowDocument, getFollowed, isUserManage, deleteForm } from './Services/Requests';
import styles from './customComponent.module.scss';
import toast, { Toaster } from 'react-hot-toast';
import { ListView, IViewField, GroupOrder, IGrouping } from "@pnp/spfx-controls-react/lib/ListView";

export interface IObjectParam {
    myProperty: string;
}

export interface ICustomComponentProps {
    pageUrlParam? : string;
    pageTitleParam? : string;
    pageFileTypeParam? : string;   
    pageId?: string;
    pageContext?: PageContext; 
    sphttpClient?: SPHttpClient;
    msGraphClientFactory?: MSGraphClientFactory;
    pages?: any;
}

export function CustomComponent (props: ICustomComponentProps){

    console.log("props.pages", props.pages);

    initializeFileTypeIcons();
    const dateOptions: any = { year: 'numeric', month: 'long', day: 'numeric' };
    
    const [myFollowedItems, setMyfollowedItems] = React.useState([]);

    React.useEffect(()=>{
        getFollowed(props.msGraphClientFactory).then(res => {
            console.log("setMyfollowedItems(res)", res);
            setMyfollowedItems(res);
        });
    }, []);
    React.useEffect(()=>{
    }, [myFollowedItems.toString()]);


    // Follow & Unfollow
    const followDocHandler = (page: any) => {
        console.log("followDocHandler", page);
        followDocument(props.msGraphClientFactory, page.SiteId, page.WebId, page.ListId, page.ListItemID).then(() => {
            setMyfollowedItems(prev => {
                const currentFollowedItems = [...prev];
                currentFollowedItems.push({name: decodeURI(page.Filename), driveId: page.DriveId});
                console.log("currentFollowedItems", currentFollowedItems);
                return currentFollowedItems;
            });
            toast.custom((t) => (
                <div className={styles.toastMsg}>
                  <Icon iconName='Accept' /> Added to <a target='_blank' href="https://www.office.com/mycontent">Favorites!</a>
                </div>
            ));
        });
    };
    const unFollowDocHandler = (page: any) => {
        console.log("followDocHandler", page);
        unFollowDocument(props.msGraphClientFactory, page.SiteId, page.WebId, page.ListId, page.ListItemID).then(()=>{
            setMyfollowedItems(prev => {
                const currentFollowedItems = prev.filter(item => !(item.name === decodeURI(page.Filename) && item.driveId === page.DriveId));
                console.log("currentFollowedItems", currentFollowedItems);
                return currentFollowedItems;
            });
            toast.custom((t) => (
                <div className={styles.toastMsg}>
                  <Icon iconName='Accept' /> Removed from <a target='_blank' href="https://www.office.com/mycontent">Favorites!</a>
                </div>
            ));
        });
    };


    const viewFields:IViewField [] = [
        {
          name: '',
          displayName: '',
          minWidth: 16,
          maxWidth: 16,
          render: (page: any) => (
            <div className={styles.favIconBtns}>
                {myFollowedItems.find(item => item.name === decodeURI(page.Filename) && item.driveId === page.DriveId ) ? 
                    <IconButton title='Unfavorite' onClick={() => unFollowDocHandler(page)} iconProps={{iconName : 'FavoriteStarFill'}} />
                : 
                    <IconButton title='Favorite' onClick={() => followDocHandler(page)} iconProps={{iconName : 'FavoriteStar'}} />
                }
                {page.FileType !== 'SharePoint.Link' &&
                    <a className={styles.attachmentLinkDownload} href={`${page.DefaultEncodingURL}`} title='Download' download>
                        <Icon iconName='Download' />
                    </a>
                } 
            </div>
          ),
        },
        {
            name: '',
            displayName: '',
            minWidth: 16,
            maxWidth: 16,
            render: (page: any) => (
                <>
                    <TooltipHost content={`${page.FileType} file`}>
                        <Icon {...getFileTypeIconProps({extension: page.FileType, size: 16}) }/>
                    </TooltipHost> 
                    <a className={styles.defautlLink + ' ' + styles.docLink} target="_blank" data-interception="off" href={page.DefaultEncodingURL}>{page.Title}</a>
                </>
            )
        },
        {
            name: 'name',
            displayName : 'Form',
            minWidth: 150,
            maxWidth: 450,
            sorting: true,
            isResizable: true,
            render : (item: any) => (
                <div>            
                    <a className={styles.defautlLink} target="_blank" data-interception="off" href={item.link}>{item.name}</a>
                </div>
            )
        },
        {
            name: 'team',
            displayName: 'Team',
            render : (page: any) => (
                <div>
                    {page.MMIntranetDeptSubDeptGrouping && page.MMIntranetDeptSubDeptGrouping.split('|')[1]}
                </div>
            )
        }
    ];
    //   const groupByFields: IGrouping[] = [
    //     {
    //         name: "deptGrp", 
    //         order: GroupOrder.ascending 
    //     },
    //     {
    //         name: "subDeptGrp", 
    //         order: GroupOrder.ascending 
    //     }
    //   ];

    return(
        <>
            <Toaster position='bottom-center' toastOptions={{custom:{duration: 4000}}}/>
            
            <div className={styles.listViewNoWrap}>
				<table className={styles.customTable} cellPadding='0' cellSpacing='0'>
                    <colgroup>
                        <col width={'10%'} />
                        <col width={'60%'} />
                        <col width={'20%'} />
                    </colgroup>
					<thead>
						<tr>
							<th></th>
							<th>Form</th>							
							<th>Team</th>							
						</tr>
					</thead>
					<tbody>
                        {props.pages.items.map(page => {
                            return (
                                <tr key={page.ListItemID}>
                                    <td>
                                        <div className={styles.formItem}>
                                            <div className={styles.favIconBtns}>
                                                {myFollowedItems.find(item => item.name === decodeURI(page.Filename) && item.driveId === page.DriveId ) ? 
                                                    <IconButton title='Unfavorite' onClick={() => unFollowDocHandler(page)} iconProps={{iconName : 'FavoriteStarFill'}} />
                                                : 
                                                    <IconButton title='Favorite' onClick={() => followDocHandler(page)} iconProps={{iconName : 'FavoriteStar'}} />
                                                }
                                            </div>
                                            <div className={styles.cellDiv}> 
                                                {page.FileType !== 'SharePoint.Link' &&
                                                    <a className={styles.attachmentLinkDownload} href={`${page.DefaultEncodingURL}`} title='Download' download>
                                                        <Icon iconName='Download' />
                                                    </a>
                                                }                                             
                                            </div>
                                        </div>
                                    </td>
                                    <td>
                                        <div className={styles.formItem}>
                                            <div className={styles.cellDiv}>
                                                <TooltipHost content={`${page.FileType} file`}>
                                                    <Icon {...getFileTypeIconProps({extension: page.FileType, size: 16}) }/>
                                                </TooltipHost> 
                                                <a className={styles.defautlLink + ' ' + styles.docLink} target="_blank" data-interception="off" href={page.DefaultEncodingURL}>{page.Title}</a>
                                            </div>
                                        </div>
                                    </td>
                                    <td>
                                        {page.MMIntranetDeptSubDeptGrouping && page.MMIntranetDeptSubDeptGrouping.split('|')[1]}
                                    </td>
                                </tr>
                            );
                        })}
                    </tbody>
				</table>
			</div>

            {/* <ListView
                items={[]}
                // viewFields={viewFields}
                // groupByFields={groupByFields}
                // compact={true}
            /> */}
            
        </>
    );

}

export class MyCustomComponentWebComponent extends BaseWebComponent {
    
    private sphttpClient: SPHttpClient;
    private pageContext: PageContext;
    private msGraphClientFactory: MSGraphClientFactory;

    public constructor() {
        super(); 
        this._serviceScope.whenFinished(()=>{
            this.pageContext = this._serviceScope.consume(PageContext.serviceKey);
            this.sphttpClient = this._serviceScope.consume(SPHttpClient.serviceKey);
            this.msGraphClientFactory = this._serviceScope.consume(MSGraphClientFactory.serviceKey);
        });
    }
 
    public async connectedCallback() {
        let props = this.resolveAttributes();
        const customComponent = <CustomComponent pageContext={this.pageContext} sphttpClient={this.sphttpClient} msGraphClientFactory={this.msGraphClientFactory} {...props}/>;
        ReactDOM.render(customComponent, this);
    }    
}
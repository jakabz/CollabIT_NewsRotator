import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";  
import { INewsItem } from './IListItem';  
  
export class ListViewService {
  
    public setup(context: any): void {  
        sp.setup({  
            spfxContext: context  
        });
    }

    public async getRotatorNews(fixedItems:boolean): Promise<INewsItem[]> {
        return new Promise<INewsItem[]>(async (resolve, reject) => {  
            try {
                const today = new Date().toISOString();
                const filter = `(NewsRotator eq 1) and (PromotedState eq 2) and (FinalApproved eq 1) and (FSObjType eq 0) and (ExpireDate gt '${today}' )`;
                const select = `ID,Title,BannerImageUrl,FileRef`;
                const top = fixedItems ? 3 : 5;
                sp.web.lists.getByTitle('Site Pages').items.filter(filter).select(select).orderBy('FirstPublishedDate', false).top(top).get().then((items:INewsItem[]) => {  
                    resolve(items);  
                });  
            }  
            catch (error) {  
                console.log(error);  
            }  
        });  
    }

    public async getAllNews(): Promise<INewsItem[]> {
        return new Promise<INewsItem[]>(async (resolve, reject) => {  
            try {
                
                const items:INewsItem[] = [];
                const getSections = async (start) => {
                    const limit = 50;
                    const section = await this.getNewsSection(start,limit);
                    items.push(...section);
                    if(section.length === limit) {
                        await getSections(section[limit-1].Id);
                    }
                }
                await getSections(0);
                resolve(items.sort(this.compare));
            }  
            catch (error) {  
                console.log(error);  
            }  
        });  
    }

    private compare( a, b ) {
        if ( a.Title < b.Title ){
          return -1;
        }
        if ( a.Title > b.Title ){
          return 1;
        }
        return 0;
    }

    private async getNewsSection(start:number,limit:number): Promise<INewsItem[]> {
        return new Promise<INewsItem[]>(async (resolve, reject) => {
            try {
                const filter = `(ID gt ${start}) and (PromotedState eq 2) and (FinalApproved eq 1) and (FSObjType eq 0)`;
                const select = `ID,Title,BannerImageUrl,FileRef`;
                const top = limit;
                sp.web.lists.getByTitle('Site Pages').items.filter(filter).select(select).orderBy('ID', true).top(top).get().then((items:INewsItem[]) => {  
                    resolve(items);  
                });
            }
            catch (error) {
                console.log(error);
            }
        })
    }

    
}  
  
const SPListViewService = new ListViewService();  
export default SPListViewService;
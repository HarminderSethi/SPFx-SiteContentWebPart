import { IColumn } from "office-ui-fabric-react/lib/DetailsList";
import { ISiteContent } from "./ISiteContent";

export interface IDetailsListContentState {
  columns: IColumn[];
  items: ISiteContent[];
  selectionDetails: string;
  isModalSelection: boolean;
  isCompactMode: boolean;
  announcedMessage?: string;
}

import { IColumn, ISelection } from "office-ui-fabric-react/lib/DetailsList";
import { IUser } from "./IUser";


export interface IMsGraphCallsState {
  columns: IColumn[];
  items: IUser[];
  selectedItems: IUser[];
  //selectionDetails: string,
}

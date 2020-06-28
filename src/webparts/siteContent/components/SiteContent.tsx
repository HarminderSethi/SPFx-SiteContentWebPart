import * as React from "react";
import styles from "./SiteContent.module.scss";

import { ISiteContentProps } from "./ISiteContentProps";
import { escape } from "@microsoft/sp-lodash-subset";

import HttpService from "../services/HttpService";
import { ISiteContent } from "./ISiteContent";
import * as _ from "lodash";
import { IDetailsListContentState } from "./IDetailsListContentState";
import { Fabric } from "office-ui-fabric-react/lib/Fabric";
import {
  DetailsList,
  DetailsListLayoutMode,
  Selection,
  SelectionMode,
  IColumn,
} from "office-ui-fabric-react/lib/DetailsList";
import { Link } from "office-ui-fabric-react/lib/Link";
import { mergeStyleSets } from "office-ui-fabric-react/lib/Styling";

const classNames = mergeStyleSets({
  fileIconHeaderIcon: {
    padding: 0,
    fontSize: "16px",
  },
  fileIconCell: {
    textAlign: "center",
    selectors: {
      "&:before": {
        content: ".",
        display: "inline-block",
        verticalAlign: "middle",
        height: "100%",
        width: "0px",
        visibility: "hidden",
      },
    },
  },
  fileIconImg: {
    verticalAlign: "middle",
    maxHeight: "16px",
    maxWidth: "16px",
  },
  controlWrapper: {
    display: "flex",
    flexWrap: "wrap",
  },
  exampleToggle: {
    display: "inline-block",
    marginBottom: "10px",
    marginRight: "30px",
  },
  selectionDetails: {
    marginBottom: "20px",
  },
});
const controlStyles = {
  root: {
    margin: "0 30px 20px 0",
    maxWidth: "300px",
  },
};
export default class SiteContent extends React.Component<
  ISiteContentProps,
  IDetailsListContentState
> {
  public constructor(props: ISiteContentProps) {
    super(props);

    const columns: IColumn[] = [
      {
        key: "ListType",
        name: "File Type",
        className: classNames.fileIconCell,
        iconClassName: classNames.fileIconHeaderIcon,
        iconName: "Page",
        isIconOnly: true,
        fieldName: "name",
        minWidth: 16,
        maxWidth: 16,
        onColumnClick: this._onColumnClick,
        onRender: (item: ISiteContent) => {
          return <img src={item.imageUrl}></img>;
        },
      },
      {
        key: "Column2",
        name: "Title",
        fieldName: "title",
        minWidth: 100,
        maxWidth: 150,
        isRowHeader: true,
        isResizable: true,
        isSorted: false,
        isSortedDescending: false,
        sortAscendingAriaLabel: "Sorted A to Z",
        sortDescendingAriaLabel: "Sorted Z to A",
        data: "string",
        isPadded: true,
        onColumnClick: this._onColumnClick,
        onRender: (item: ISiteContent) => {
          return (
            <Link href={item.url} target="_blank">
              {item.title}
            </Link>
          );
        },
      },
      {
        key: "Column3",
        name: "Created",
        fieldName: "createdDate",
        minWidth: 100,
        maxWidth: 150,
        isRowHeader: true,
        isResizable: true,
        isSorted: false,
        isSortedDescending: false,
        sortAscendingAriaLabel: "Sorted A to Z",
        sortDescendingAriaLabel: "Sorted Z to A",
        data: "string",
        isPadded: true,
        onColumnClick: this._onColumnClick,
      },
      {
        key: "Column4",
        name: "Modified",
        fieldName: "lastModifiedDate",
        minWidth: 100,
        maxWidth: 150,
        isRowHeader: true,
        isResizable: true,
        isSorted: false,
        isSortedDescending: false,
        sortAscendingAriaLabel: "Sorted A to Z",
        sortDescendingAriaLabel: "Sorted Z to A",
        data: "string",
        isPadded: true,
        onColumnClick: this._onColumnClick,
      },
      {
        key: "Column5",
        name: "Items",
        fieldName: "itemCount",
        minWidth: 100,
        maxWidth: 150,
        isRowHeader: true,
        isResizable: true,
        isSorted: false,
        isSortedDescending: false,
        sortAscendingAriaLabel: "Sorted A to Z",
        sortDescendingAriaLabel: "Sorted Z to A",
        data: "string",
        isPadded: true,
        onColumnClick: this._onColumnClick,
      },
    ];

    this.state = {
      items: [],
      columns: columns,
      selectionDetails: "",
      isModalSelection: false,
      isCompactMode: false,
    };
  }
  public render(): React.ReactElement<ISiteContentProps> {
    const { columns, isCompactMode, items } = this.state;
    return (
      <Fabric>
        <DetailsList
          items={items}
          compact={isCompactMode}
          columns={columns}
          selectionMode={SelectionMode.none}
          layoutMode={DetailsListLayoutMode.justified}
          isHeaderVisible={true}
          selectionPreservedOnEmptyClick={true}
          onItemInvoked={this._onItemInvoked}
        />
      </Fabric>
    );
  }

  private _onItemInvoked(item: ISiteContent): void {
    if (item.url != undefined && item.url != "") {
      window.open(item.url, "_blank");
    }
  }

  public componentDidUpdate(prevProps) {
    if (
      prevProps.viewSiteContentBy != this.props.viewSiteContentBy ||
      prevProps.orderBy != this.props.orderBy
    ) {
      const contentArray: ISiteContent[] = [];
      HttpService.GetSiteContent(this.props)
        .then((data) => {
          if (data.value != undefined) {
            data.value.forEach((element) => {
              contentArray.push({
                id: element.Id,
                title: element.Title,
                url: element.RootFolder.ServerRelativeUrl,
                itemCount: element.ItemCount,
                lastModifiedDate: this.getFormattedDate(
                  element.LastItemModifiedDate
                ),
                createdDate: this.getFormattedDate(element.Created),
                imageUrl: this.props.siteUrl + "/" + element.ImageUrl,
                entityTypeName: element.EntityTypeName,
              });
            });
            this.updateState(contentArray);
          }
        })
        .catch((error) => console.log(error));
    }
  }

  public componentDidMount() {
    const contentArray: ISiteContent[] = [];
    HttpService.GetSiteContent(this.props)
      .then((data) => {
        if (data.value != undefined) {
          data.value.forEach((element) => {
            contentArray.push({
              id: element.Id,
              title: element.Title,
              url: element.RootFolder.ServerRelativeUrl,
              itemCount: element.ItemCount,
              lastModifiedDate: this.getFormattedDate(
                element.LastItemModifiedDate
              ),
              createdDate: this.getFormattedDate(element.Created),
              imageUrl: this.props.siteUrl + "/" + element.ImageUrl,
              entityTypeName: element.EntityTypeName,
            });
          });
          this.updateState(contentArray);
        }
      })
      .catch((error) => console.log(error));
  }
  private getFormattedDate(inputDate: string): string {
    let formattedDate = "";
    if (inputDate != "") {
      let date = new Date(inputDate);
      formattedDate =
        date.getMonth() + 1 + "/" + date.getDate() + "/" + date.getFullYear();
    }
    return formattedDate;
  }

  private _onColumnClick = (
    ev: React.MouseEvent<HTMLElement>,
    column: IColumn
  ): void => {
    const { columns, items } = this.state;
    const newColums: IColumn[] = columns.slice();
    const currColumn: IColumn = newColums.filter(
      (currCol) => column.key === currCol.key
    )[0];
    newColums.forEach((newCol: IColumn) => {
      if (newCol === currColumn) {
        currColumn.isSortedDescending = !currColumn.isSortedDescending;
        currColumn.isSorted = true;
      } else {
        newCol.isSorted = false;
        newCol.isSortedDescending = true;
      }
    });
    const newItems = this._copyAndSort(
      items,
      currColumn.fieldName!,
      currColumn.isSortedDescending
    );
    this.setState({
      columns: newColums,
      items: newItems,
    });
  };

  private _copyAndSort<T>(
    items: T[],
    columnKey: string,
    isSortedDescending?: boolean
  ): T[] {
    const key = columnKey as keyof T;
    return items
      .slice(0)
      .sort((a: T, b: T) =>
        (isSortedDescending ? a[key] < b[key] : a[key] > b[key]) ? 1 : -1
      );
  }

  private updateState(contentArray: ISiteContent[]): void {
    let columnName: string;
    let isSortDesc: boolean = false;
    if (this.props.orderBy == undefined || this.props.orderBy == "") {
      columnName = "Created";
      isSortDesc = true;
    } else if (this.props.orderBy != "") {
      if (this.props.orderBy.toLowerCase() == "modifieddesc") {
        columnName = "Modified";
        isSortDesc = true;
      } else if (this.props.orderBy.toLowerCase() == "modifiedasc") {
        columnName = "Modified";
        isSortDesc = false;
      } else if (this.props.orderBy.toLowerCase() == "createddesc") {
        columnName = "Created";
        isSortDesc = true;
      } else if (this.props.orderBy.toLowerCase() == "createdasc") {
        columnName = "Created";
        isSortDesc = false;
      } else if (this.props.orderBy.toLowerCase() == "titledesc") {
        columnName = "Title";
        isSortDesc = true;
      } else if (this.props.orderBy.toLowerCase() == "titleasc") {
        columnName = "Title";
        isSortDesc = false;
      }
    }
    const newColums: IColumn[] = this.state.columns.slice();
    const currColumn: IColumn = newColums.filter(
      (currCol) => currCol.name === columnName
    )[0];
    newColums.forEach((newCol: IColumn) => {
      if (newCol === currColumn) {
        currColumn.isSortedDescending = isSortDesc;
        currColumn.isSorted = true;
      } else {
        newCol.isSorted = false;
        newCol.isSortedDescending = true;
      }
    });
    const newItems = this._copyAndSort(
      contentArray,
      currColumn.fieldName!,
      currColumn.isSortedDescending
    );
    this.setState({
      columns: newColums,
      items: newItems,
    });
  }
}

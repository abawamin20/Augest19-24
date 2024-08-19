import * as React from "react";
import {
  DetailsList,
  DetailsListLayoutMode,
  SelectionMode,
  IColumn,
  DetailsHeader,
  Selection,
  IDetailsListStyles,
} from "@fluentui/react/lib/DetailsList";
import { mergeStyles } from "@fluentui/react";
import "./styles.css";
import { IColumnInfo } from "../PagesList/PagesService";
import { WebPartContext } from "@microsoft/sp-webpart-base";

const gridStyles: Partial<IDetailsListStyles> = {
  root: {},
  headerWrapper: {},
  contentWrapper: {
    maxHeight: "600px",
    scrollbarColor: "#f5f5f5",
    scrollbarWidth: "thin",
    overflowX: "hidden",
    overflowY: "auto",
    width: "max-content",
  },
};

const customHeaderClass = mergeStyles({
  backgroundColor: "#efefef", // Custom header background color
  color: "white", // Custom text color
  paddingTop: 0,
  paddingBottom: 0,
  selectors: {
    "& .ms-DetailsHeader": {
      backgroundColor: "#0078d4", // Custom header background color
      borderBottom: "1px solid #ccc",
    },
  },
});

export interface IReusableDetailListcomponents {
  columns: (
    columns: IColumnInfo[],
    context: WebPartContext,
    currentUser: any,
    onColumnClick: any,
    sortBy: string,
    isDecending: boolean,
    setShowFilter: (column: IColumn, columnType: string) => void
  ) => IColumn[];
  columnInfos: IColumnInfo[];
  currentUser: any;
  context: WebPartContext;
  setShowFilter: (column: IColumn, columnType: string) => void;
  updateSelection: (selection: Selection) => void;
  items: any[];
  sortPages: (column: IColumn, isAscending: boolean) => void;
  sortBy: string;
  siteUrl: string;
  isDecending: boolean;
}

export class ReusableDetailList extends React.Component<
  IReusableDetailListcomponents,
  {}
> {
  private _selection: Selection;
  constructor(components: IReusableDetailListcomponents) {
    super(components);

    this._selection = new Selection({
      onSelectionChanged: () => {
        this.props.updateSelection(this._selection);
      },
      getKey: this._getKey,
    });
  }

  componentDidMount() {
    window.addEventListener("contentLoaded", this.handleContentLoaded);
  }

  componentWillUnmount() {
    window.removeEventListener("contentLoaded", this.handleContentLoaded);
  }

  componentDidUpdate(prevcomponents: IReusableDetailListcomponents) {
    if (prevcomponents.items !== this.props.items) {
      this.forceUpdate();
      window.dispatchEvent(new Event("contentLoaded"));
    }
  }

  handleContentLoaded = () => {
    const navSection: HTMLElement | null =
      document.querySelector(".custom-nav");
    const detailSection: HTMLElement | null =
      document.querySelector(".detail-display");

    function adjustNavHeight() {
      if (navSection && detailSection) {
        const detailHeight = detailSection.offsetHeight;
        const minHeight = 500; // Minimum height in pixels
        navSection.style.height = `${Math.max(detailHeight, minHeight)}px`;
      }
    }

    adjustNavHeight();
    window.addEventListener("resize", adjustNavHeight);
  };
  _onRenderDetailsHeader = (components: any) => {
    if (!components) {
      return null;
    }

    // Apply custom styles to the header
    return (
      <DetailsHeader
        {...components}
        className="stickyHeader"
        styles={{
          root: customHeaderClass, // Apply custom styles
        }}
      />
    );
  };

  public render() {
    const {
      columnInfos,
      currentUser,
      context,
      columns,
      items,
      sortPages,
      sortBy,
      isDecending,
      setShowFilter,
    } = this.props;

    return (
      <div>
        <DetailsList
          styles={gridStyles}
          items={items}
          compact={true}
          columns={columns(
            columnInfos,
            context,
            currentUser,
            sortPages,
            sortBy,
            isDecending,
            setShowFilter
          )}
          selectionMode={SelectionMode.single}
          selection={this._selection}
          getKey={this._getKey}
          setKey="none"
          layoutMode={DetailsListLayoutMode.justified}
          isHeaderVisible={true}
          onRenderDetailsHeader={this._onRenderDetailsHeader}
          onItemInvoked={this._onItemInvoked}
          className="detailList"
        />
      </div>
    );
  }
  private _getKey(item: any, index?: number): string {
    return item.key;
  }

  private _onItemInvoked = (item: any): void => {
    window.open(`${this.props.siteUrl}${item.FileRef}`, "_blank");
  };
}

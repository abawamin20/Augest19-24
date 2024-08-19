import * as React from "react";
import { Dialog, DialogFooter } from "@fluentui/react";
import { SPHttpClient } from "@microsoft/sp-http";
import { ReusableDetailList } from "../common/ReusableDetailList";
import PagesService, { FilterDetail, IColumnInfo } from "./PagesService";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { PagesColumns } from "./PagesColumns";
import { DefaultButton, IColumn, Icon, Selection } from "@fluentui/react";
import { makeStyles, useId, Input } from "@fluentui/react-components";
import styles from "./pages.module.scss";
import "./pages.css";
import { FilterPanelComponent } from "./PanelComponent";
// import AlertForm from "../Forms/AlertMeForm";

export interface IPagesListProps {
  context: WebPartContext;
  selectedViewId: string;
  feedbackPageUrl: string;
}

const useStyles = makeStyles({
  root: {
    display: "flex",
    gap: "2px",
    maxWidth: "400px",
    alignItems: "center",
  },
});

const PagesList = (props: IPagesListProps) => {
  const subscribeLink: string = "/_layouts/15/SubNew.aspx";
  const alertLink: string = "/_layouts/15/mySubs.aspx";

  // Destructure the props
  const { context, selectedViewId } = props;

  /**
   * State variables for the component.
   */

  // Options for the page size dropdown
  const [pageSizeOption] = React.useState<number[]>([
    10, 15, 20, 40, 60, 80, 100,
  ]);

  const [hideFeedBackDialog, setHideFeedBackDialog] = React.useState(true);

  const toggleHideFeedbackDialog = () => {
    setHideFeedBackDialog(!hideFeedBackDialog);
  };
  const [hideAlertMeDialog, setHideAlertMeDialog] = React.useState(true);
  const [hideManageAlertDialog, setHideManageAlertDialog] =
    React.useState(true);

  const toggleHideAlertMeDialog = () => {
    setHideAlertMeDialog(!hideAlertMeDialog);
  };
  const toggleHideManageAlertDialog = () => {
    setHideManageAlertDialog(!hideAlertMeDialog);
  };

  const [columnInfos, setColumnInfos] = React.useState<IColumnInfo[]>([]);

  // The search text for filtering pages
  const [searchText, setSearchText] = React.useState<string>(""); // Initially set to empty string

  // The list of pages
  const [pages, setPages] = React.useState<any[]>([]); // Initially set to empty array

  // The selected category
  const [catagory, setCatagory] = React.useState<string | null>(null); // Initially set to empty string
  const [isLoading, setIsLoading] = React.useState<boolean>(false); // Initially set to empty string

  // The initial list of pages
  const [initialPages, setInitialPages] = React.useState<any[]>([]); // Initially set to empty array

  // The paginated list of pages
  const [paginatedPages, setPaginatedPages] = React.useState<any[]>([]); // Initially set to empty array

  // The column to sort by
  const [sortBy, setSortBy] = React.useState<string>(""); // Initially set to empty string

  // The current page number
  const [currentPageNumber, setCurrentPageNumber] = React.useState<number>(1); // Initially set to 1

  // The total number of pages
  const [totalPages, setTotalPages] = React.useState<number>(1); // Initially set to 1

  // The number of items to display per page
  const [pageSize, setPageSize] = React.useState<number>(10); // Initially set to 10

  // The total number of items
  const [totalItems, setTotalItems] = React.useState<number>(0); // Initially set to 0

  // The sorting order
  const [isDecending, setIsDecending] = React.useState<boolean>(false); // Initially set to false

  // Whether to show the filter panel
  const [showFilter, setShowFilter] = React.useState<boolean>(false); // Initially set to false

  // The column to filter by
  const [filterColumn, setFilterColumn] = React.useState<string>(""); // Initially set to empty string

  // The type of column to filter by
  const [filterColumnType, setFilterColumnType] = React.useState<string>(""); // Initially set to empty string

  // The filter details
  const [filterDetails, setFilterDetails] = React.useState<FilterDetail[]>([]); // Initially set to empty array

  // The taxonomy filter details
  const [taxonomyFilters, setTaxonomyFilters] = React.useState<FilterDetail[]>(
    []
  );

  // The filter details
  const [selectionDetails, setSelectionDetails] = React.useState<any | []>([]);
  // The filter details
  const [listId, setListId] = React.useState<string>("");
  const [currentUser, setCurrentUser] = React.useState<any>(null);
  const [viewId, setViewId] = React.useState<string>("");

  // Create an instance of the PagesService class
  const pagesService = new PagesService(context);

  // Get a unique id for the input field
  const inputId = useId("input");

  // Get the styles for the input field
  const inputStyles = useStyles();

  const subscribeIframeRef = React.useRef<HTMLIFrameElement>(null);
  const feedBackIframeRef = React.useRef<HTMLIFrameElement>(null);

  React.useEffect(() => {
    const checkIframeUrl = () => {
      const iframe = subscribeIframeRef.current;
      if (iframe && iframe.contentWindow) {
        const currentUrl = iframe.contentWindow.location.href;

        if (
          currentUrl.indexOf("blank") === -1 &&
          currentUrl.indexOf(subscribeLink) === -1
        ) {
          setHideAlertMeDialog(true);

          getPages(catagory, columnInfos);
        }
      }
    };

    // Check the URL every 2 seconds
    const intervalId = setInterval(checkIframeUrl, 2000);

    // Clean up the interval on component unmount
    return () => clearInterval(intervalId);
  }, [subscribeLink, setHideAlertMeDialog, catagory, catagory]);

  React.useEffect(() => {
    const checkIframeUrl = () => {
      const iframe = feedBackIframeRef.current;
      if (iframe && iframe.contentWindow) {
        const currentUrl = iframe.contentWindow.location.href;
        console.log(currentUrl);
        if (
          currentUrl.indexOf("blank") === -1 &&
          currentUrl.indexOf(props.feedbackPageUrl) === -1
        ) {
          setHideFeedBackDialog(true);
        }
      }
    };

    // Check the URL every 2 seconds
    const intervalId = setInterval(checkIframeUrl, 2000);

    // Clean up the interval on component unmount
    return () => clearInterval(intervalId);
  }, [props.feedbackPageUrl, setHideFeedBackDialog]);
  /**
   * Resets the filters by clearing the checked items and
   * calling the applyFilters function with an empty filter detail.
   */
  const resetFilters = () => {
    // Clear the filter details
    setFilterDetails([]);
    setTaxonomyFilters([]);

    // Clear the search text
    setSearchText("");

    // Call the fetchPages function with the default arguments
    fetchPages(1, pageSize, "Created", true, "", catagory, []);
  };

  /**
   * Fetches the paginated pages based on the given parameters.
   *
   * @param {number} [page=1] - The page number to fetch. Defaults to 1.
   * @param {number} [pageSizeAmount=pageSize] - The number of items per page. Defaults to the `pageSize` state variable.
   * @param {string} [sortBy="Created"] - The column to sort by. Defaults to "Created".
   * @param {boolean} [isSortedDescending=isDecending] - Whether to sort in descending order. Defaults to the `isDecending` state variable.
   * @param {string} [searchText=""] - The search text to filter by. Defaults to an empty string.
   * @param {string} [category=catagory] - The category to filter by. Defaults to the `catagory` state variable.
   * @param {FilterDetail[]} filterDetails - The filter details to apply.
   *
   * @return {Promise<void>} - A promise that resolves when the paginated pages are fetched.
   */
  const fetchPages = (
    page = 1,
    pageSizeAmount = pageSize,
    sortBy = "Created",
    isSortedDescending = isDecending,
    searchText = "",
    category = catagory,
    filterDetails: FilterDetail[],
    taxonomyFilters: FilterDetail[] = [],
    columns: IColumnInfo[] = columnInfos
  ) => {
    const url = `${context.pageContext.web.serverRelativeUrl}/SitePages/${category}`;

    return (
      pagesService
        // Call the pagesService to fetch the filtered pages with non-taxonomy filters
        .getFilteredPages(
          page,
          pageSizeAmount,
          sortBy,
          isSortedDescending,
          url,
          searchText,
          filterDetails, // Send only non-taxonomy filters to API
          columns
        )
        .then((res) => {
          // If there are taxonomy filters, apply them locally
          if (taxonomyFilters.length > 0) {
            // Apply the taxonomy filters locally on the API results
            res = res.filter((item) => {
              // For each item, ensure that it matches every taxonomy filter
              return taxonomyFilters.every((taxonomyFilter) => {
                // Check if the item has the taxonomy field and if it is an array
                if (Array.isArray(item[taxonomyFilter.filterColumn])) {
                  // Ensure that the item matches at least one of the values in the filter
                  return taxonomyFilter.values.some((value) => {
                    // Check if the value matches any of the Label in the taxonomy items
                    return item[taxonomyFilter.filterColumn].some(
                      (taxonomyItem: { Label: string; TermGuid: string }) =>
                        taxonomyItem.Label === value
                    );
                  });
                }
                return false; // If the field does not exist or is not an array, filter it out
              });
            });
          }

          // Set the total number of items in the filtered response
          setTotalItems(res.length);

          // Calculate the total number of pages based on the page size
          const totalPages = Math.ceil(res.length / pageSizeAmount);

          // If there are no pages, set the total number of pages to 1
          if (totalPages === 0) {
            setTotalPages(1);
          } else {
            // Otherwise, set the total number of pages based on the page size
            setTotalPages(totalPages);
          }

          // Slice the response to get the paginated pages for the current page number
          const paginatedData = res.slice(0, pageSizeAmount);

          // Set the state with paginated pages
          setPaginatedPages(paginatedData);

          // Set the state with all the pages (full filtered result)
          setPages(res);

          // Return the full filtered response
          return res;
        })
    );
  };

  /**
   * Fetches the pages from the given path and filter categories
   * and updates the state with the initial pages
   * @param path - The path to the SitePages library
   */
  const getPages = async (
    path: string | null,
    columns: IColumnInfo[]
  ): Promise<void> => {
    setIsLoading(true);
    // Get the initial pages from the API
    const initialPagesFromApi = await fetchPages(
      1,
      pageSize,
      "Created",
      true,
      searchText,
      path,
      filterDetails,
      [],
      columns
    );

    // Update the state with the initial pages
    setInitialPages(initialPagesFromApi);
    setIsLoading(false);
  };

  /**
   * Applies the given filter details to filter the pages
   *
   * @param {FilterDetail} filterDetail - The filter detail object containing the filter details
   */
  const applyFilters = (filterDetail: FilterDetail): void => {
    /**
     * Updates the current filter details state with the new filter detail,
     * or removes the filter detail if the values array is empty.
     *
     */
    let currentFilters: FilterDetail[] = filterDetails;
    let currentTaxonomyFilters: FilterDetail[] = taxonomyFilters;

    if (filterDetail.filterColumnType === "TaxonomyFieldTypeMulti") {
      if (filterDetail.values.length === 0) {
        currentTaxonomyFilters = taxonomyFilters.filter(
          (item) => item.filterColumn !== filterDetail.filterColumn
        );
      } else {
        currentTaxonomyFilters = [
          ...taxonomyFilters.filter(
            (item) => item.filterColumn !== filterDetail.filterColumn
          ),
          filterDetail,
        ];
      }
    } else {
      if (filterDetail.values.length === 0) {
        currentFilters = filterDetails.filter(
          (item) => item.filterColumn !== filterDetail.filterColumn
        );
      } else
        currentFilters = [
          ...filterDetails.filter(
            (item) => item.filterColumn !== filterDetail.filterColumn
          ),
          filterDetail,
        ];
    }

    setFilterDetails(currentFilters);
    setTaxonomyFilters(currentTaxonomyFilters);

    fetchPages(
      1, // Page number
      pageSize, // Page size
      "Created", // Sorting criteria
      true, // Sorting order (ascending/descending)
      searchText, // Search text
      catagory, // Category (assuming this is another state or prop)
      currentFilters, // Updated filter details,
      currentTaxonomyFilters
    );
  };

  /**
   * Sort the pages list based on the specified column.
   *
   * @param {IColumn} column - The column to sort by.
   */
  const sortPages = (column: IColumn) => {
    // Set the sort by column state
    setSortBy(column.fieldName as string);

    // If the column is the same as the current sort by column, toggle the sort order
    if (column.fieldName === sortBy) {
      setIsDecending(!isDecending);
    } else {
      // Otherwise, set the sort order to descending
      setIsDecending(true);
    }

    // Fetch the pages list with the new sort criteria
    fetchPages(
      1, // Page number
      pageSize, // Page size
      column.fieldName, // Sorting criteria
      column.isSortedDescending, // Sorting order (ascending/descending)
      searchText, // Search text
      catagory, // Category (assuming this is another state or prop)
      filterDetails, // Filter details
      taxonomyFilters
    );
  };

  const handlePageChange = (page: number, pageSizeChanged = pageSize) => {
    // Ensure page is an integer
    const currentPage = Math.ceil(page);

    // Update current page number state
    setCurrentPageNumber(currentPage);

    // Calculate slice indices for pagination
    const startIndex = (currentPage - 1) * pageSizeChanged;

    const endIndex = currentPage * pageSizeChanged;

    // Slice the 'pages' array to get the current page of data
    const paginated = pages.slice(startIndex, endIndex);

    setTotalPages(Math.ceil(totalItems / pageSizeChanged));
    // Update paginated pages state
    setPaginatedPages(paginated);
  };

  /**
   * Handles the search functionality by fetching pages with specified parameters.
   */
  const handleSearch = () => {
    fetchPages(
      1, // Page number
      pageSize, // Page size
      "Created", // Sorting criteria
      true, // Sorting order (ascending/descending)
      searchText, // Search text
      catagory, // Category
      filterDetails, // Filter details
      taxonomyFilters
    );
  };

  /**
   * Navigates to the first page of paginated data.
   */
  const goToFirstPage = () => handlePageChange(1);

  /**
   * Navigates to the last page of paginated data.
   */
  const goToLastPage = () => handlePageChange(totalPages);

  /**
   * Navigates to the previous page of paginated data.
   * If the current page is the first page, no action is taken.
   */
  const goToPreviousPage = () =>
    handlePageChange(Math.max(currentPageNumber - 1, 1));

  /**
   * Navigates to the next page of paginated data.
   * If the current page is the last page, no action is taken.
   */
  const goToNextPage = () =>
    handlePageChange(Math.min(currentPageNumber + 1, totalPages));

  /**
   * Handles the input change event.
   * Parses the input value to an integer and calls handlePageChange with the parsed value.
   * If the input is not a number, calls handlePageChange with 0.
   *
   * @param e - The event object
   */
  const handleInputChange = (e: any) => {
    const inputValue = e.target.value;

    if (!isNaN(inputValue)) {
      const page = parseInt(inputValue, 10);
      handlePageChange(page);
    } else {
      handlePageChange(0);
    }
  };

  /**
   * Handles the change event of the page size dropdown.
   *
   * This function is triggered when the user selects a new page size from the dropdown.
   * It updates the page size state and calls the `handlePageChange` function to update
   * the paginated data.
   *
   * @function handlePageSizeChange
   * @memberof PagesList
   *
   * @param {any} e - The event object.
   * @return {void}
   */
  const handlePageSizeChange = (e: any) => {
    // Update the page size state
    setPageSize(e.target.value);
    // Handle the page change with the new page size
    handlePageChange(1, e.target.value);
  };

  /**
   * Dismisses the filter panel.
   * Sets the showFilter state to false.
   *
   * @function dismissPanel
   * @memberof PagesList
   * @returns {void}
   */
  const dismissPanel = (): void => {
    setShowFilter(false);
  };

  const getColumns = async (selectedViewId: string) => {
    const columns = await pagesService.getColumns(selectedViewId);

    setColumnInfos(columns);

    return columns;
  };

  React.useEffect(() => {
    const handleEvent = (e: any) => {
      const selectedCategory = e.detail;
      setCatagory(selectedCategory);
      getColumns(selectedViewId).then((col) => {
        getPages(selectedCategory, col);
      });
    };

    const fetchCurrentUser = async () => {
      try {
        const currentUserResponse = await context.spHttpClient.get(
          `${context.pageContext.web.absoluteUrl}/_api/web/currentuser`,
          SPHttpClient.configurations.v1
        );
        const userData = await currentUserResponse.json();
        setCurrentUser(userData);
      } catch (error) {
        console.error("Error fetching current user:", error);
      }
    };
    fetchCurrentUser();
    pagesService.getListDetailsByName("Site Pages").then((res) => {
      setListId(res.Id);
    });

    window.addEventListener("catagorySelected", handleEvent);
  }, []);

  React.useEffect(() => {
    if (viewId !== selectedViewId) {
      setViewId(selectedViewId);
      getColumns(selectedViewId).then((col) => {
        getPages(catagory, col);
      });
    }
  }, [selectedViewId]);

  return (
    <div className="w-pageSize0 detail-display">
      {showFilter && (
        <FilterPanelComponent
          isOpen={showFilter}
          headerText="Filter Articles"
          applyFilters={applyFilters}
          dismissPanel={dismissPanel}
          selectedItems={
            [...filterDetails, ...taxonomyFilters].filter(
              (item) => item.filterColumn === filterColumn
            )[0] || { filterColumn: "", values: [] }
          }
          columnName={filterColumn}
          columnType={filterColumnType}
          pagesService={pagesService}
          data={initialPages}
        />
      )}
      <div className={`${styles.top}`}>
        <div
          className={`${styles["first-section"]} d-flex justify-content-between align-items-end py-2 px-2`}
        >
          <span className={`fs-4 ${styles["knowledgeText"]}`}>
            {catagory && <span className="">{catagory}</span>}
          </span>
          <div className={`${inputStyles.root} d-flex align-items-center me-2`}>
            <Input
              id={inputId}
              value={searchText}
              onChange={(e) => setSearchText(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === "Enter") {
                  handleSearch();
                }
              }}
              placeholder="Search"
            />
          </div>
        </div>

        <div
          className={`d-flex justify-content-between align-items-center fs-5 px-2 my-2`}
        >
          <span>Articles /</span>
          {totalItems > 0 ? (
            <div className="d-flex align-items-center">
              {selectionDetails && selectionDetails.length > 0 && (
                <DefaultButton
                  className="me-2"
                  onClick={() => {
                    toggleHideAlertMeDialog();
                  }}
                >
                  <span className="d-flex align-items-center">
                    <Icon iconName="Ringer" className="me-2" />
                    Alert Me
                  </span>
                </DefaultButton>
              )}
              {selectionDetails && selectionDetails.length > 0 && (
                <DefaultButton
                  className="me-2"
                  onClick={() => {
                    toggleHideManageAlertDialog();
                  }}
                >
                  <span className="d-flex align-items-center">
                    <Icon iconName="EditNote" className="me-2" />
                    Manage My Alerts
                  </span>
                </DefaultButton>
              )}
              {selectionDetails && selectionDetails.length > 0 && (
                <DefaultButton
                  className="me-2"
                  onClick={() => {
                    toggleHideFeedbackDialog();
                  }}
                >
                  <span className="d-flex align-items-center">
                    <Icon iconName="Feedback" className="me-2" />
                    Add Feedback
                  </span>
                </DefaultButton>
              )}
              {((filterDetails && filterDetails.length > 0) ||
                (taxonomyFilters && taxonomyFilters.length > 0)) && (
                <DefaultButton
                  onClick={() => {
                    resetFilters();
                  }}
                >
                  Clear
                </DefaultButton>
              )}
              <span className="ms-2 fs-6">Results ({totalItems})</span>
            </div>
          ) : (
            <span className="fs-6">No articles to display</span>
          )}
        </div>
      </div>

      {!isLoading && (
        <ReusableDetailList
          items={paginatedPages}
          context={context}
          columns={PagesColumns}
          columnInfos={columnInfos}
          currentUser={currentUser}
          setShowFilter={(column: IColumn, columnType: string) => {
            setShowFilter(!showFilter);
            setFilterColumn(column.fieldName as string);
            setFilterColumnType(columnType);
          }}
          updateSelection={(selection: Selection) => {
            setSelectionDetails(selection.getSelection());
          }}
          sortPages={sortPages}
          sortBy={sortBy}
          siteUrl={window.location.origin}
          isDecending={isDecending}
        />
      )}
      <div className="d-flex justify-content-end">
        <div
          className="d-flex align-items-center my-1"
          style={{
            fontSize: "13px",
          }}
        >
          <div className="d-flex align-items-center me-3">
            <span className={`me-2 ${styles.blueText}`}>Items / Page </span>
            <select
              className="form-select"
              value={pageSize}
              onChange={handlePageSizeChange}
              name="pageSize"
              style={{
                width: 80,
                height: 35,
              }}
            >
              {pageSizeOption.map((pageSize) => {
                return (
                  <option key={pageSize} value={pageSize}>
                    {pageSize}
                  </option>
                );
              })}
            </select>
          </div>
          <span className={`me-2 ${styles.blueText}`}>Page</span>
          <input
            type="text"
            value={currentPageNumber}
            onChange={handleInputChange}
            className="form-control"
            style={{
              width: 50,
              height: 35,
            }}
          />
          <span className="fs-6 mx-2">of {totalPages}</span>
          <span
            onClick={goToFirstPage}
            className={`mx-2 ${styles["pagination-btns"]} ${
              currentPageNumber === 1 && styles.disabledPagination
            }`}
          >
            <i className="fa fa-step-backward" aria-hidden="true"></i>
          </span>
          <span
            onClick={goToPreviousPage}
            className={`mx-2 ${styles["pagination-btns"]} ${
              currentPageNumber === 1 && styles.disabledPagination
            }`}
          >
            <i className="fa fa-caret-left" aria-hidden="true"></i>
          </span>
          <span
            onClick={goToNextPage}
            className={`mx-2 ${styles["pagination-btns"]} ${
              currentPageNumber >= totalPages ? styles.disabledPagination : ""
            }`}
          >
            <i className="fa fa-caret-right" aria-hidden="true"></i>
          </span>
          <span
            onClick={goToLastPage}
            className={`mx-2 ${styles["pagination-btns"]} ${
              currentPageNumber >= totalPages ? styles.disabledPagination : ""
            }`}
          >
            <i className="fa fa-step-forward" aria-hidden="true"></i>
          </span>
        </div>
      </div>

      <Dialog
        hidden={hideFeedBackDialog}
        onDismiss={toggleHideFeedbackDialog}
        modalProps={{
          isBlocking: false,
        }}
        maxWidth="90vw"
        minWidth="60vw"
      >
        <iframe
          src={`${props.feedbackPageUrl}`}
          width="100%"
          height="600px"
          style={{ border: "none" }}
          ref={feedBackIframeRef}
        ></iframe>

        <DialogFooter>
          <DefaultButton
            onClick={() => setHideFeedBackDialog(true)}
            text="Close"
          />
        </DialogFooter>
      </Dialog>
      <Dialog
        hidden={hideAlertMeDialog}
        onDismiss={toggleHideAlertMeDialog}
        modalProps={{
          isBlocking: false,
        }}
        maxWidth="90vw"
        minWidth="60vw"
      >
        <iframe
          ref={subscribeIframeRef}
          src={`${
            context.pageContext.web.absoluteUrl
          }${subscribeLink}?List=${listId}&Id=${
            selectionDetails[0] && selectionDetails[0].Id
          }`}
          width="100%"
          height="600px"
          style={{ border: "none" }}
        ></iframe>

        <DialogFooter>
          <DefaultButton
            onClick={() => setHideAlertMeDialog(true)}
            text="Close"
          />
        </DialogFooter>
      </Dialog>
      <Dialog
        hidden={hideManageAlertDialog}
        onDismiss={toggleHideManageAlertDialog}
        modalProps={{
          isBlocking: false,
        }}
        maxWidth="90vw"
        minWidth="60vw"
      >
        <iframe
          src={`${context.pageContext.web.absoluteUrl}${alertLink}`}
          width="100%"
          height="600px"
          style={{ border: "none" }}
          id="alertFrame"
        ></iframe>

        <DialogFooter>
          <DefaultButton
            onClick={() => setHideManageAlertDialog(true)}
            text="Close"
          />
        </DialogFooter>
      </Dialog>
    </div>
  );
};

export default PagesList;

import * as React from "react";
import type { IPagesDisplayProps } from "./IPagesDisplayProps";
import PagesList from "./PagesList/PagesList";
import { SPComponentLoader } from "@microsoft/sp-loader";

interface IPagesDisplayState {
  spLoaded: boolean;
}

export default class PagesDisplay extends React.Component<
  IPagesDisplayProps,
  IPagesDisplayState
> {
  constructor(props: IPagesDisplayProps) {
    super(props);
    this.state = {
      spLoaded: false,
    };

    // Load CSS files
    const cssURLs = [
      "https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/css/bootstrap.min.css",
      "https://maxcdn.bootstrapcdn.com/font-awesome/4.6.3/css/font-awesome.min.css",
    ];
    cssURLs.forEach((url) => SPComponentLoader.loadCss(url));

    // Initialize SharePoint scripts
    this.loadSPScripts();
  }

  private loadSPScripts(): void {
    Promise.all([
      SPComponentLoader.loadScript("/_layouts/15/init.js", {
        globalExportsName: "$_global_init",
      }),
      SPComponentLoader.loadScript("/_layouts/15/MicrosoftAjax.js", {
        globalExportsName: "Sys",
      }),
      SPComponentLoader.loadScript(
        "/_layouts/15/ScriptResx.ashx?name=sp.res&culture=en-us",
        { globalExportsName: "Sys" }
      ),
      SPComponentLoader.loadScript("/_layouts/15/SP.Runtime.js", {
        globalExportsName: "SP",
      }),
      SPComponentLoader.loadScript("/_layouts/15/SP.js", {
        globalExportsName: "SP",
      }),
      SPComponentLoader.loadScript("/_layouts/15/sp.init.js", {
        globalExportsName: "SP",
      }),
      SPComponentLoader.loadScript("/_layouts/15/sp.ui.dialog.js", {
        globalExportsName: "SP",
      }),
    ])
      .then(() => {
        console.log(SP);
        // Verify SP is available
        if (typeof SP !== "undefined") {
          this.setState({ spLoaded: true });
        } else {
          console.error("SP is not defined after loading scripts.");
        }
      })
      .catch((error) => {
        console.error("Error loading SharePoint scripts:", error);
      });
  }

  public render(): React.ReactElement<IPagesDisplayProps> {
    return (
      <React.Fragment>
        {this.state.spLoaded ? (
          <PagesList
            context={this.props.context}
            selectedViewId={this.props.selectedViewId}
            feedbackPageUrl={this.props.feedbackPageUrl}
          />
        ) : (
          <div>Loading SharePoint scripts...</div>
        )}
      </React.Fragment>
    );
  }
}

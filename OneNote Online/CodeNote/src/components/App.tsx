import * as React from "react";
import { Button, ButtonType } from "office-ui-fabric-react";
import Splash from "./Splash";
// import HeroList, { HeroListItem } from "./HeroList";
import Progress from "./Progress";

import * as OfficeHelpers from "@microsoft/office-js-helpers";

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

export interface AppState {
  //   listItems: HeroListItem[];
}

export default class App extends React.Component<AppProps, AppState> {
  constructor(props, context) {
    super(props, context);
    this.state = {
      listItems: []
    };
  }

  componentDidMount() {
    this.setState({
      listItems: [
        {
          icon: "Ribbon",
          primaryText: "Achieve more with Office integration"
        },
        {
          icon: "Unlock",
          primaryText: "Unlock features and functionality"
        },
        {
          icon: "Design",
          primaryText: "Create and visualize like a pro"
        }
      ]
    });
  }

  insertCode = async () => {
    try {
      await OneNote.run(async context => {
        /**
         * Insert your OneNote code here
         */
        var code = (document.getElementById("codeBox") as HTMLInputElement)
          .value;

        var lines = code.split(/\r|\r\n|\n/);
        var rows = lines.length;

        // Get the current page.
        var page = context.application.getActivePage();

        // Get the current outline.
        var activeOutline = context.application.getActiveOutlineOrNull();

        console.log(activeOutline);

        // // Queue a command to load the page with the title property.
        page.load("contents");

        // Creates a 2D array with the content.
        var table = new Array(rows);

        for (var i = 0; i < table.length; i++) {
          table[i] = new Array(2);
          table[i][0] = (i + 1).toString();
          table[i][1] = lines[i];
        }

        // Add text to the page by using the specified HTML.
        if (activeOutline != null) {
          let codeTable: OneNote.Table = activeOutline.appendTable(1, 2);

          // Creates the row numbers.
          var lineNumbers = new Array(rows);
          for (var i = 0; i < lineNumbers.length; i++) {
            lineNumbers[i] = new Array(1);
            lineNumbers[i][0] = (i + 1).toString();
          }

          let lineNumbersTable: OneNote.Table = codeTable
            .getCell(0, 0)
            .appendTable(rows, 1, lineNumbers);
          lineNumbersTable.borderVisible = false;
        } else {
          page
            .addOutline(40, 80, "<p></p>")
            .appendTable(rows, 2, table).borderVisible = false;
        }

        // Run the queued commands, and return a promise to indicate task completion.
        return context
          .sync()
          .then(function() {
            //console.log("Added outline to page ", page.contents);
          })
          .catch(function(error) {
            // App.showNotification("Error: " + error);
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
              console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
          });
      });
    } catch (error) {
      OfficeHelpers.UI.notify(error);
      OfficeHelpers.Utilities.log(error);
    }
  };

  render() {
    const { title, isOfficeInitialized } = this.props;

    if (!isOfficeInitialized) {
      return (
        <Progress
          title={title}
          logo="assets/CodeNoteText.png"
          message="Please sideload your addin to see app body."
        />
      );
    }

    return (
      <div className="splash-screen">
        <Splash logo="assets/CodeNoteText.png" title={this.props.title} />
        <div>
          <p>
            Language:
            <select>
              <option value="C#">C#</option>
              <option value="C++">C++</option>
              <option value="Java">Java</option>
              <option value="JavaScript">JavaScript</option>
              <option value="PHP">PHP</option>
            </select>
          </p>
          <p>
            Theme:
            <select>
              <option value="BlendDark">Blend (Dark)</option>
              <option value="Atom">Atom</option>
            </select>
          </p>
        </div>
        <br />
        <textarea className="codebox" id="codeBox" />
        <Button buttonType={ButtonType.command} onClick={this.insertCode}>
          Insert Code
        </Button>
        {/* <HeroList
          message="Discover what CodeNote can do for you today!"
          items={this.state.listItems}
        >
          <p className="ms-font-l">
            Modify the source files, then click <b>Run</b>.
          </p>
          <Button
            className="ms-welcome__action"
            buttonType={ButtonType.hero}
            iconProps={{ iconName: "ChevronRight" }}
            onClick={this.click}
          >
            Run
          </Button>
        </HeroList> */}
      </div>
    );
  }
}

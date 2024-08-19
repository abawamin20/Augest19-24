import * as React from "react";
import { useState } from "react";
import { TextField, Dropdown, Checkbox, PrimaryButton } from "@fluentui/react";
import PagesService from "../PagesList/PagesService";

export interface IAlertFormProps {
  pageService: PagesService;
  pageTitle: string;
  pageId: number;
  listId: string;
  currentUser: any;
  toggleHideFeedbackDialog: () => void;
}
const AlertForm: React.FunctionComponent<IAlertFormProps> = (props) => {
  const [alertTitle, setAlertTitle] = useState<string>(
    `Site Pages: ${props.pageTitle}`
  );
  const [emailNotification, setEmailNotification] = useState<boolean>(false);
  const [frequency, setFrequency] = useState<string>("Immediate");

  const frequencyOptions = [
    { key: "Immediate", text: "Immediate" },
    { key: "Daily", text: "Daily" },
    { key: "Weekly", text: "Weekly" },
  ];

  const handleSubmit = async () => {
    SP.SOD.executeFunc("sp.js", "SP.ClientContext", function () {
      var clientContext = SP.ClientContext.get_current();
      var web = clientContext.get_web();
      var list = web.get_lists().getByTitle("Site Pages"); // Adjust list title if needed

      // Create alert
      var alertCreationData = {
        AlertFrequency: 0, // Immediate
        AlertType: 1, // Modified items
        EventType: 1, // All changes
        List: props.listId,
        ItemId: 12, // Item ID of the target item
        Title: "Alert Title",
        UserId: 11, // User ID for the alert
        AlwaysNotify: true,
        Status: 1, // Active
      };

      var alertCreationRequest = new SP.AlertCreationRequest(alertCreationData);

      clientContext.load(list);
      clientContext.executeQueryAsync(function () {
        alertCreationRequest.executeQueryAsync(
          function () {
            console.log("Alert created successfully.");
          },
          function (sender: any, args: any) {
            console.error(
              "Failed to create alert. Error: " + args.get_message()
            );
          }
        );
      });
    });

    // await props.pageService.saveUserAlert(
    //   alertTitle,
    //   props.listId,
    //   props.pageId,
    //   frequency,
    //   props.currentUser.Id
    // );
    // try {
    //   // Initialize the SharePoint context
    //   const ctx = SP.ClientContext.get_current();

    //   // Get the list by title
    //   const list = ctx.get_web().get_lists().getByTitle("Site Pages"); // Adjust the list title if needed

    //   // Define CAML query to get the specific item by FileLeafRef
    //   const camlQuery = new SP.CamlQuery();
    //   camlQuery.set_viewXml(`
    //   <View>
    //   <Query>
    //   <Where>
    //     <Eq>
    //       <FieldRef Name='FileLeafRef' />
    //       <Value Type='Text'>${props.pageTitle}</Value>
    //     </Eq>
    //   </Where>
    //   </Query>
    //   </View>
    //   `);

    //   // Get items based on the query
    //   const listItems = list.getItems(camlQuery);
    //   ctx.load(listItems);
    //   ctx.executeQueryAsync(
    //     () => {
    //       const itemsEnumerator = listItems.getEnumerator();
    //       let found = false;
    //       while (itemsEnumerator.moveNext()) {
    //         const listItem = itemsEnumerator.get_current();
    //         const filePath = listItem.get_item("FileRef"); // Full path of the file

    //         if (filePath.includes(props.parentFolderPath)) {
    //           console.log("File found in the correct folder:", filePath);
    //           found = true;
    //           // Perform operations with the found item
    //           // For example, you might want to call another function to process this file

    //           createAlert(ctx, listItem);

    //           break;
    //         }
    //       }
    //       if (!found) {
    //         console.log("No file found in the specified folder.");
    //       }
    //     },
    //     (sender: any, args: any) => {
    //       console.log(sender);
    //       // On failure
    //       console.error(
    //         "Failed to retrieve list items. Error: ",
    //         args.get_message()
    //       );
    //     }
    //   );
    // } catch (error) {
    //   console.error("Error creating alert:", error);
    // }
  };

  // const createAlert = (ctx: any, listItem: any) => {
  //   try {
  //     const user = ctx.get_web().get_currentUser();
  //     const alerts = user.get_alerts();

  //     const alertTime = new Date();
  //     alertTime.setHours(alertTime.getHours() + 24);

  //     const notify = new SP.AlertCreationInformation();
  //     notify.set_title(alertTitle);
  //     notify.set_alertFrequency(SP.AlertFrequency[frequency]); // Ensure frequency is set correctly
  //     notify.set_alertType(SP.AlertType.list); // This is usually list, but we are going to filter
  //     notify.set_list(listItem); // Specify list for alerting
  //     notify.set_deliveryChannels(SP.AlertDeliveryChannel.email);
  //     notify.set_alwaysNotify(true);
  //     notify.set_status(SP.AlertStatus.on);
  //     notify.set_alertTime(alertTime);
  //     notify.set_user(user);
  //     notify.set_eventType(SP.AlertEventType.all);

  //     // Assuming 'Filter' is used for additional specificity if supported
  //     const freq = props.pageService
  //       .getAlertFrequencyValue(frequency)
  //       .toString();
  //     notify.set_filter(freq);

  //     alerts.add(notify);
  //     user.update();

  //     ctx.executeQueryAsync(
  //       () => {
  //         console.log("Alert created successfully.");
  //         props.toggleHideFeedbackDialog();
  //       },
  //       (sender: any, args: any) => {
  //         console.log(sender);
  //         console.error("Failed to create alert. Error: ", args.get_message());
  //       }
  //     );
  //   } catch (error) {
  //     console.error("Error creating alert:", error);
  //   }
  // };

  return (
    <div>
      <h2>Create Alert</h2>

      <TextField
        label="Alert Title"
        value={alertTitle}
        onChange={(_, value) => setAlertTitle(value || "")}
        style={{
          marginBottom: "10px",
        }}
      />

      <Dropdown
        label="Alert Frequency"
        selectedKey={frequency}
        options={frequencyOptions}
        onChange={(_, option) => setFrequency(option?.key as string)}
        style={{
          marginBottom: "10px",
        }}
      />

      <Checkbox
        label="Email Notification"
        checked={emailNotification}
        onChange={(_, checked) => setEmailNotification(checked || false)}
      />

      <PrimaryButton
        style={{
          marginTop: "10px",
        }}
        text="Create Alert"
        onClick={handleSubmit}
      />
    </div>
  );
};

export default AlertForm;

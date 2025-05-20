import { IInputs, IOutputs } from "../generated/ManifestTypes";
import {
  DetailsList,
  DetailsListLayoutMode,
  Selection,
  IColumn,
  mergeStyleSets,
  TextField,
  IButtonStyles,
  Dropdown,
  IDropdownOption,
  PrimaryButton,
  DefaultButton,
  CheckboxVisibility,
  Link,
} from "@fluentui/react";
import * as React from "react";
import CustomEmailForm, { EmailFormProps } from "./Forms/CustomEmailForm";
import TaskForm, { TaskFormProps } from "./Forms/TaskForm";
import PhoneCallForm, { PhoneCallFormProps } from "./Forms/PhoneCallForm";
import AppointmentForm, { AppointmentFormProps } from "./Forms/AppointmentForm";
import { emailData } from "./Backend/EmailData";
import { taskData } from "./Backend/TaskData";
import { phoneCallData } from "./Backend/PhoneCallData";
import { appointmentData } from "./Backend/AppointmentData";

// Interface for combined activity records
interface IRecord {
  id: string;
  subject: string;
  regardingobjectidname: string;
  createdon: string;
  status: string;
  priority: string;
  scheduledend: string;
  activitytypecode: "Email" | "Task" | "PhoneCall" | "Appointment";
  createdbyname: string;
  createdby: string;
  [key: string]: string | number | boolean | null | undefined | IAttachment[] | IPartyList | IPartyList[] | undefined;
}

// Interface for attachment details (for emails)
interface IAttachment {
  attachmentId: string;
  body: string;
  filename: string;
  filesize: number;
  mimetype: string;
}

interface IPartyList {
  partyid: string;
  name: string;
  entitytype: string;
}

export interface IActivityDatasetControlProps {
  context: ComponentFramework.Context<IInputs>;
  notifyOutputChanged: () => void;
}

export interface IActivityDatasetControlState {
  items: IRecord[];
  filteredItems: IRecord[];
  columns: IColumn[];
  selectedItems: IRecord[];
  searchQuery: string;
  activityTypeFilter: string;
  isEmailFormOpen: boolean;
  isTaskFormOpen: boolean;
  isPhoneCallFormOpen: boolean;
  isAppointmentFormOpen: boolean;
  selectedRecord: IRecord | null;
  sortColumn: string | null;
  isSortDescending: boolean;
  isLoading: boolean;
  errorMessage: string | null;
  currentPage: number;
  pageSize: number;
}

const classNames = mergeStyleSets({
  headerWrapper: {
    overflow: "hidden",
    position: "sticky",
    top: 0,
    zIndex: 1,
    backgroundColor: "lightblue",
  },
  container: {
    height: "300px",
    overflow: "auto",
    position: "relative",
    border: "0px solid black",
  },
  pagination: {
    display: "flex",
    justifyContent: "center",
    alignItems: "center",
    marginTop: "10px",
  },
  subjectLink: {
    color: "#0078d4",
    textDecoration: "none",
    ":hover": {
      textDecoration: "underline",
      cursor: "pointer",
    },
  },
});

const buttonStyles: IButtonStyles = {
  root: {
    border: "none",
    backgroundColor: "transparent",
    padding: "0 4px",
    selectors: {
      ":hover": { backgroundColor: "transparent" },
      ":active": { backgroundColor: "transparent" },
    },
  },
  icon: { color: "inherit" },
  label: { color: "inherit" },
};

const activityTypeOptions: IDropdownOption[] = [
  { key: "All", text: "All" },
  { key: "Email", text: "Email" },
  { key: "Task", text: "Task" },
  { key: "PhoneCall", text: "Phone Call" },
  { key: "Appointment", text: "Appointment" },
];

export class ActivityDatasetControl extends React.Component<
  IActivityDatasetControlProps,
  IActivityDatasetControlState
> {
  private _selection: Selection;

  constructor(props: IActivityDatasetControlProps) {
    super(props);

    this.state = {
      items: [],
      filteredItems: [],
      columns: this.getColumns(),
      selectedItems: [],
      searchQuery: "",
      activityTypeFilter: "All",
      isEmailFormOpen: false,
      isTaskFormOpen: false,
      isPhoneCallFormOpen: false,
      isAppointmentFormOpen: false,
      selectedRecord: null,
      sortColumn: "createdon",
      isSortDescending: true,
      isLoading: false,
      errorMessage: null,
      currentPage: 1,
      pageSize: 5,
    };

    this._selection = new Selection({
      onSelectionChanged: () => {
        this.setState({ selectedItems: this._selection.getSelection() as IRecord[] });
      },
    });
  }

  private getColumns(): IColumn[] {
    return [
      {
        key: "subject",
        name: "Subject",
        fieldName: "subject",
        minWidth: 100,
        maxWidth: 180,
        isResizable: true,
        headerClassName: classNames.headerWrapper,
        isSorted: false,
        isSortedDescending: false,
        onColumnClick: this.onColumnClick,
        onRender: (item: IRecord) => (
          <Link
            className={classNames.subjectLink}
            onClick={() => this.openRecord(item)}
          >
            {item.subject || "N/A"}
          </Link>
        ),
      },
      {
        key: "createdon",
        name: "Created On",
        fieldName: "createdon",
        minWidth: 100,
        maxWidth: 200,
        isResizable: true,
        headerClassName: classNames.headerWrapper,
        isSorted: true,
        isSortedDescending: true,
        onColumnClick: this.onColumnClick,
        onRender: (item: IRecord) => item.createdon || "N/A",
      },
      {
        key: "activitytypecode",
        name: "Activity Type",
        fieldName: "activitytypecode",
        minWidth: 100,
        maxWidth: 150,
        isResizable: true,
        headerClassName: classNames.headerWrapper,
        isSorted: false,
        isSortedDescending: false,
        onColumnClick: this.onColumnClick,
        onRender: (item: IRecord) => item.activitytypecode || "N/A",
      },
      {
        key: "status",
        name: "Activity Status",
        fieldName: "status",
        minWidth: 100,
        maxWidth: 150,
        isResizable: true,
        headerClassName: classNames.headerWrapper,
        isSorted: false,
        isSortedDescending: false,
        onColumnClick: this.onColumnClick,
        onRender: (item: IRecord) => item.status || "N/A",
      },
      {
        key: "scheduledend",
        name: "Due Date",
        fieldName: "scheduledend",
        minWidth: 100,
        maxWidth: 150,
        isResizable: true,
        headerClassName: classNames.headerWrapper,
        isSorted: false,
        isSortedDescending: false,
        onColumnClick: this.onColumnClick,
        onRender: (item: IRecord) => item.scheduledend || "N/A",
      },
      {
        key: "createdby",
        name: "Created By",
        fieldName: "createdby",
        minWidth: 100,
        maxWidth: 150,
        isResizable: true,
        headerClassName: classNames.headerWrapper,
        isSorted: false,
        isSortedDescending: false,
        onColumnClick: this.onColumnClick,
        onRender: (item: IRecord) => (
          <Link
            className={classNames.subjectLink}
            onClick={() =>
              item.createdby && item.createdby_entitytype
                ? this.openCrmForm(String(item.createdby_entitytype), String(item.createdby))
                : null
            }
            disabled={!item.createdby || !item.createdby_entitytype}
          >
            {item.createdbyname || "N/A"}
          </Link>
        ),
      },
      {
        key: "regarding",
        name: "Regarding",
        fieldName: "regarding",
        minWidth: 100,
        maxWidth: 150,
        isResizable: true,
        headerClassName: classNames.headerWrapper,
        isSorted: false,
        isSortedDescending: false,
        onColumnClick: this.onColumnClick,
        onRender: (item: IRecord) => item.regarding || "N/A",
      },
    ];
  }

  private async restoreRecords(): Promise<void> {

    alert("Restore functionality is not implemented yet.");

    //#region Functionality To Restore Records in CRM
    // const { selectedItems } = this.state;
    // const { context } = this.props;

    // if (selectedItems.length === 0) {
    //   this.setState({ errorMessage: "No records selected for restoration." });
    //   return;
    // }

    // this.setState({ isLoading: true, errorMessage: null });

    // try {
    //   for (const record of selectedItems) {
    //     let entityName: string;
    //     let entityData: any = {
    //       subject: record.subject,
    //       regardingobjectid: record.regardingobjectid ? {
    //         "@odata.id": `${record.regardingobjectid_entitytype}(${record.regardingobjectid})`
    //       } : undefined,
    //       createdon: record.createdon,
    //       statuscode: record.status,
    //       prioritycode: record.priority,
    //       scheduledend: record.scheduledend,
    //       ownerid: record.ownerid ? {
    //         "@odata.id": `${record.ownerid_entitytype}(${record.ownerid})`
    //       } : undefined,
    //     };

    //     switch (record.activitytypecode) {
    //       case "Email":
    //         entityName = "email";
    //         entityData = {
    //           ...entityData,
    //           description: record.description,
    //           torecipients: record.torecipients,
    //           from: record.from,
    //           directioncode: record.directioncode === "true",
    //         };
    //         break;
    //       case "Task":
    //         entityName = "task";
    //         entityData = {
    //           ...entityData,
    //           description: record.description,
    //           strava_url: record.strava_url,
    //           strava_contact: record.strava_contact,
    //           strava_homenumber: record.strava_homenumber,
    //           strava_mobile: record.strava_mobile,
    //         };
    //         break;
    //       case "PhoneCall":
    //         entityName = "phonecall";
    //         entityData = {
    //           ...entityData,
    //           phonenumber: record.phonenumber,
    //           directioncode: record.directioncode === "true",
    //           new_phonecallreason: record.new_phonecallreason,
    //           new_phonecalloutcome: record.new_phonecalloutcome,
    //         };
    //         break;
    //       case "Appointment":
    //         entityName = "appointment";
    //         entityData = {
    //           ...entityData,
    //           scheduledstart: record.scheduledstart,
    //           new_appttype: record.new_appttype,
    //           new_apptmethod: record.new_apptmethod,
    //           location: record.location,
    //         };
    //         break;
    //       default:
    //         throw new Error(`Unsupported activity type: ${record.activitytypecode}`);
    //     }

    //     // Remove undefined or null values
    //     Object.keys(entityData).forEach(key => 
    //       (entityData[key] === undefined || entityData[key] === null) && delete entityData[key]
    //     );

    //     // Create record using Web API
    //     await context.webAPI.createRecord(entityName, entityData);
    //   }

    //   // Refresh the view after successful restoration
    //   await this.updateView();
    //   this.setState({
    //     isLoading: false,
    //     errorMessage: "Records restored successfully.",
    //     selectedItems: [],
    //   });
    //   this._selection.setAllSelected(false);

    //   // Clear success message after 3 seconds
    //   setTimeout(() => {
    //     this.setState({ errorMessage: null });
    //   }, 3000);

    // } catch (error) {
    //   console.error("Error restoring records:", error);
    //   this.setState({
    //     isLoading: false,
    //     errorMessage: `Failed to restore records: ${error instanceof Error ? error.message : "Unknown error"}`,
    //   });
    // }
    //#endregion
  }

  openCrmForm = (entityName: string, id: string) => {
    const orgUrl = "https://nsminc-sandbox.crm.dynamics.com";
    const formUrl = `${orgUrl}/main.aspx?etn=${encodeURIComponent(entityName)}&id=${encodeURIComponent(id)}&pagetype=entityrecord`;
    try {
      window.open(formUrl, "_blank");
      console.log(`Opened ${entityName} form in a new tab successfully.`);
    } catch (error) {
      const errMsg = (error instanceof Error) ? error.message : String(error);
      console.error(`Error opening ${entityName} form in a new tab: ${errMsg}`);
    }
  };

  componentDidMount() {
    this.updateView();
  }

  componentDidUpdate(prevProps: IActivityDatasetControlProps, prevState: IActivityDatasetControlState) {
    if (
      prevProps.context.parameters.recordId.raw !== this.props.context.parameters.recordId.raw ||
      prevState.activityTypeFilter !== this.state.activityTypeFilter ||
      prevState.searchQuery !== this.state.searchQuery ||
      prevState.currentPage !== this.state.currentPage
    ) {
      console.log(
        `componentDidUpdate triggered: recordId=${this.props.context.parameters.recordId.raw}, ` +
        `activityTypeFilter=${this.state.activityTypeFilter}, searchQuery=${this.state.searchQuery}, ` +
        `currentPage=${this.state.currentPage}`
      );
      this.updateView();
    }
  }

  async updateView(): Promise<void> {
    this.setState({ isLoading: true, errorMessage: null });

    const recordId = "af9540d7-9ff5-ef11-be20-7c1e5232f746";

    if (!recordId) {
      this.setState({
        items: [],
        filteredItems: [],
        isLoading: false,
        errorMessage: "No parent record ID provided.",
      });
      return;
    }

    try {
      const items: IRecord[] = [
        ...emailData.map((email) => ({
          ...email,
          activitytypecode: "Email" as const,
          regarding: email.regardingobjectidname || "",
          regardingobjectidname: email.regardingobjectidname || "",
          createdon: email.createdon || "",
          status: String(email.status ?? ""),
          priority: String(email.priority ?? ""),
          scheduledend: email.scheduledend || "",
          createdby: email.createdby || "N/A",
        })),
        ...taskData.map((task) => ({
          ...task,
          activitytypecode: "Task" as const,
          regarding: task.regardingobjectidname || "",
          regardingobjectidname: task.regardingobjectidname || "",
          createdon: task.createdon || "",
          status: String(task.status ?? ""),
          priority: String(task.priority ?? ""),
          scheduledend: task.scheduledend || "",
          createdby: task.createdby || "N/A",
        })),
        ...phoneCallData.map((phoneCall) => ({
          ...phoneCall,
          activitytypecode: "PhoneCall" as const,
          regarding: phoneCall.regardingobjectidname || "",
          regardingobjectidname: phoneCall.regardingobjectidname || "",
          createdon: phoneCall.createdon || "",
          status: String(phoneCall.status ?? ""),
          priority: String(phoneCall.priority ?? ""),
          scheduledend: phoneCall.scheduledend || "",
          createdby: phoneCall.createdby || "N/A",
        })),
        ...appointmentData.map((appointment) => ({
          ...appointment,
          activitytypecode: "Appointment" as const,
          regarding: appointment.regardingobjectidname || "",
          regardingobjectidname: appointment.regardingobjectidname || "",
          createdon: appointment.createdon || "",
          status: String(appointment.status ?? ""),
          priority: String(appointment.priority ?? ""),
          scheduledend: appointment.scheduledend || "",
          createdby: appointment.createdby || "N/A",
        })),
      ];

      const filteredItems = this.applyFilters(items);

      console.log(`updateView: items=${items.length}, filteredItems=${filteredItems.length}`);

      this.setState({
        items,
        filteredItems,
        isLoading: false,
        errorMessage: items.length === 0 ? "No records found for this ID." : null,
      });
    } catch (error) {
      console.error("Error processing static data:", error);
      this.setState({
        items: [],
        filteredItems: [],
        isLoading: false,
        errorMessage: "Failed to load static data.",
      });
    }
  }

  applyFilters(items: IRecord[]): IRecord[] {
    const { searchQuery, activityTypeFilter, sortColumn, isSortDescending } = this.state;

    let filteredItems = [...items];

    // Apply activity type filter
    if (activityTypeFilter !== "All") {
      filteredItems = filteredItems.filter(
        (item) => item.activitytypecode === activityTypeFilter
      );
    }

    // Apply search filter only on grid columns
    if (searchQuery) {
      const columnFields = this.state.columns.map((col) => col.fieldName);
      filteredItems = filteredItems.filter((item) =>
        columnFields.some((field) =>
          field !== undefined &&
          item[field as keyof IRecord] &&
          typeof item[field as keyof IRecord] === "string" &&
          (item[field as keyof IRecord] as string).toLowerCase().includes(searchQuery.toLowerCase())
        )
      );
    }

    // Apply sorting
    const sortedItems = this.sortItems(filteredItems, sortColumn, isSortDescending);
    console.log(
      `applyFilters: activityType=${activityTypeFilter}, searchQuery=${searchQuery}, ` +
      `itemsCount=${sortedItems.length}`
    );
    return [...sortedItems];
  }

  onColumnClick = (ev: React.MouseEvent<HTMLElement>, column: IColumn): void => {
    const { columns, filteredItems, sortColumn, isSortDescending } = this.state;
    let newSortDescending = isSortDescending;

    if (sortColumn === column.key) {
      newSortDescending = !isSortDescending;
    } else {
      newSortDescending = false;
    }

    const sortedItems = this.sortItems(filteredItems, column.key, newSortDescending);
    const newColumns = columns.map((col) => ({
      ...col,
      isSorted: col.key === column.key,
      isSortedDescending: col.key === column.key && newSortDescending,
    }));

    this.setState({
      filteredItems: [...sortedItems],
      columns: newColumns,
      sortColumn: column.key,
      isSortDescending: newSortDescending,
      currentPage: 1,
    });
  };

  sortItems = (items: IRecord[], sortColumn: string | null, isSortDescending: boolean): IRecord[] => {
    if (!sortColumn) return [...items];

    return [...items].sort((a, b) => {
      const aValue = a[sortColumn] || "";
      const bValue = b[sortColumn] || "";

      if (sortColumn === "createdon" || sortColumn === "scheduledend") {
        const aDate =
          typeof aValue === "string" || typeof aValue === "number"
            ? new Date(aValue).getTime()
            : new Date("1970-01-01").getTime();
        const bDate =
          typeof bValue === "string" || typeof bValue === "number"
            ? new Date(bValue).getTime()
            : new Date("1970-01-01").getTime();
        return isSortDescending ? bDate - aDate : aDate - bDate;
      }

      const comparison = aValue.toString().localeCompare(bValue.toString());
      return isSortDescending ? -comparison : comparison;
    });
  };

  handleSearch = (
    event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
    newValue?: string
  ): void => {
    const searchQuery = newValue || "";
    this.setState((prevState) => ({
      searchQuery,
      filteredItems: this.applyFilters(prevState.items),
      currentPage: 1,
    }));
  };

  handleActivityTypeChange = (
    event: React.FormEvent<HTMLDivElement>,
    option?: IDropdownOption
  ): void => {
    const activityTypeFilter = option ? option.key.toString() : "All";
    console.log(`handleActivityTypeChange: new filter=${activityTypeFilter}`);
    this.setState((prevState) => ({
      activityTypeFilter,
      filteredItems: this.applyFilters(prevState.items),
      currentPage: 1,
    }));
  };

  handlePageChange = (page: number): void => {
    this.setState({ currentPage: page });
  };

  openRecord = (item: IRecord): void => {
    if (!item?.id || !item?.subject) {
      console.error("Invalid record selected:", item);
      return;
    }
    this.setState({
      selectedRecord: item,
      isEmailFormOpen: item.activitytypecode === "Email",
      isTaskFormOpen: item.activitytypecode === "Task",
      isPhoneCallFormOpen: item.activitytypecode === "PhoneCall",
      isAppointmentFormOpen: item.activitytypecode === "Appointment",
    });
  };

  closeForm = (): void => {
    this.setState({
      isEmailFormOpen: false,
      isTaskFormOpen: false,
      isPhoneCallFormOpen: false,
      isAppointmentFormOpen: false,
      selectedRecord: null,
    });
  };

  render() {
    const {
      filteredItems,
      columns,
      searchQuery,
      activityTypeFilter,
      isEmailFormOpen,
      isTaskFormOpen,
      isPhoneCallFormOpen,
      isAppointmentFormOpen,
      selectedRecord,
      isLoading,
      errorMessage,
      currentPage,
      pageSize,
    } = this.state;

    // Calculate pagination
    const totalItems = filteredItems.length;
    const totalPages = Math.ceil(totalItems / pageSize);
    const startIndex = (currentPage - 1) * pageSize;
    const endIndex = startIndex + pageSize;
    const paginatedItems = filteredItems.slice(startIndex, endIndex);

    return (
      <div>
        <div
          style={{
            display: "flex",
            justifyContent: "space-between",
            alignItems: "center",
            marginBottom: 10,
          }}
        >
          <div style={{ display: "flex", flexDirection: "column", gap: "5px", alignItems: "flex-start" }}>
            <label
              htmlFor="activityTypeDropdown"
              style={{ textAlign: "left" }}
            >
              Activity Type
            </label>
            <Dropdown
              id="activityTypeDropdown"
              placeholder="Select Activity Type"
              options={activityTypeOptions}
              selectedKey={activityTypeFilter}
              onChange={this.handleActivityTypeChange}
              style={{ width: "200px" }}
            />
          </div>
          <div style={{ display: "flex", gap: "10px", alignItems: "center" }}>
            <PrimaryButton
              text="Restore"
              onClick={() => this.restoreRecords()}
              disabled={this.state.selectedItems.length === 0}
            />
            <TextField
              placeholder="Search..."
              value={searchQuery}
              onChange={this.handleSearch}
              style={{ width: "200px" }}
            />
          </div>
        </div>
        {isLoading ? (
          <div>Loading...</div>
        ) : errorMessage ? (
          <div style={{ color: errorMessage.includes("successfully") ? "green" : "red" }}>
            {errorMessage}
          </div>
        ) : (
          <div>
            <div className={classNames.container}>
              <DetailsList
                key={`${activityTypeFilter}-${searchQuery}-${currentPage}`}
                items={paginatedItems}
                columns={columns}
                setKey="set"
                layoutMode={DetailsListLayoutMode.justified}
                selection={this._selection}
                selectionPreservedOnEmptyClick={true}
                onItemInvoked={this.openRecord}
                checkboxVisibility={CheckboxVisibility.always}
              />
            </div>
            {totalItems > 0 && (
              <div className={classNames.pagination}>
                <DefaultButton
                  text="Previous"
                  onClick={() => this.handlePageChange(currentPage - 1)}
                  disabled={currentPage === 1}
                  style={{ marginRight: "10px" }}
                />
                <span>
                  Page {currentPage} of {totalPages}
                </span>
                <DefaultButton
                  text="Next"
                  onClick={() => this.handlePageChange(currentPage + 1)}
                  disabled={currentPage === totalPages}
                  style={{ marginLeft: "10px" }}
                />
              </div>
            )}
          </div>
        )}
        {isEmailFormOpen && selectedRecord && (
          <CustomEmailForm
            data={{
              id: selectedRecord.id,
              subject: selectedRecord.subject,
              regarding: selectedRecord.regardingobjectidname,
              regardingobjectid_entitytype: String(selectedRecord.regardingobjectid_entitytype || ""),
              regardingobjectid: String(selectedRecord.regardingobjectid || ""),
              createdon: selectedRecord.createdon,
              ownerid: String(selectedRecord.ownerid || ""),
              ownerid_entitytype: String(selectedRecord.ownerid_entitytype || ""),
              owneridname: String(selectedRecord.owneridname || ""),
              status: selectedRecord.status,
              priority: selectedRecord.priority,
              scheduledend: selectedRecord.scheduledend,
              activitytypecode: selectedRecord.activitytypecode as "Email",
              directioncode: String(selectedRecord.directioncode || "N/A"),
              sendermailboxidname: String(selectedRecord.sendermailboxidname || "N/A"),
              torecipients: String(selectedRecord.torecipients || "N/A"),
              from: String(selectedRecord.from || "N/A"),
              to: Array.isArray(selectedRecord.to) ? selectedRecord.to as IPartyList[] : [],
              cc: Array.isArray(selectedRecord.cc) ? selectedRecord.cc as IPartyList[] : [],
              bcc: Array.isArray(selectedRecord.bcc) ? selectedRecord.bcc as IPartyList[] : [],
              senton: String(selectedRecord.senton || "N/A"),
              description: String(selectedRecord.description || ""),
              attachments: Array.isArray(selectedRecord.attachments) ? selectedRecord.attachments as IAttachment[] : [],
            }}
            onClose={this.closeForm}
            context={this.props.context}
          />
        )}
        {isTaskFormOpen && selectedRecord && (
          <TaskForm
            data={{
              id: selectedRecord.id,
              subject: selectedRecord.subject,
              regarding: selectedRecord.regardingobjectidname,
              regardingobjectid_entitytype: String(selectedRecord.regardingobjectid_entitytype || ""),
              regardingobjectid: String(selectedRecord.regardingobjectid || ""),
              createdon: selectedRecord.createdon,
              status: selectedRecord.status,
              priority: selectedRecord.priority,
              scheduledend: selectedRecord.scheduledend,
              activitytypecode: selectedRecord.activitytypecode as "Task",
              strava_url: String(selectedRecord.strava_url || ""),
              ownerid: String(selectedRecord.ownerid || ""),
              ownerid_entitytype: String(selectedRecord.ownerid_entitytype || ""),
              owneridname: String(selectedRecord.owneridname || ""),
              strava_contact: String(selectedRecord.strava_contact || ""),
              strava_contactname: String(selectedRecord.strava_contactname || ""),
              strava_contact_entitytype: String(selectedRecord.strava_contact_entitytype || ""),
              strava_homenumber: String(selectedRecord.strava_homenumber || ""),
              strava_mobile: String(selectedRecord.strava_mobile || ""),
              strava_business: String(selectedRecord.strava_business || ""),
              strava_businessdirect: String(selectedRecord.strava_businessdirect || ""),
              strava_otherphone: String(selectedRecord.strava_otherphone || ""),
              strava_connectedcallstatus: String(selectedRecord.strava_connectedcallstatus || ""),
              strava_notecategory: String(selectedRecord.strava_notecategory || ""),
              strava_note: String(selectedRecord.strava_note || ""),
              description: String(selectedRecord.description || ""),
              strava_createoutlookcalendarreminder: String(selectedRecord.strava_createoutlookcalendarreminder || ""),
              strava_publishnote: Boolean(selectedRecord.strava_publishnote),
              new_taskpriority: String(selectedRecord.new_taskpriority || ""),
              new_teamassociation: String(selectedRecord.new_teamassociation || ""),
              new_quoteboundindicator: String(selectedRecord.new_quoteboundindicator || ""),
              actualdurationminutes: String(selectedRecord.actualdurationminutes || ""),
            }}
            onClose={this.closeForm}
          />
        )}
        {isPhoneCallFormOpen && selectedRecord && (
          <PhoneCallForm
            data={{
              id: selectedRecord.id,
              subject: selectedRecord.subject,
              regarding: selectedRecord.regardingobjectidname,
              regardingobjectid_entitytype: String(selectedRecord.regardingobjectid_entitytype || ""),
              regardingobjectid: String(selectedRecord.regardingobjectid || ""),
              createdon: selectedRecord.createdon,
              status: selectedRecord.status,
              priority: selectedRecord.priority,
              scheduledend: selectedRecord.scheduledend,
              activitytypecode: selectedRecord.activitytypecode as "PhoneCall",
              phonenumber: String(selectedRecord.phonenumber || ""),
              from: String(selectedRecord.from || "N/A"),
              from_entitytype: String(selectedRecord.from_entitytype || ""),
              fromname: String(selectedRecord.fromname || ""),
              to: Array.isArray(selectedRecord.to) ? selectedRecord.to as IPartyList[] : [],
              to_entitytype: String(selectedRecord.to_entitytype || ""),
              toname: String(selectedRecord.toname || ""),
              directioncode: String(selectedRecord.directioncode || ""),
              scheduledstart: String(selectedRecord.scheduledstart || ""),
              senton: String(selectedRecord.senton || ""),
              actualstart: String(selectedRecord.actualstart || ""),
              actualend: String(selectedRecord.actualend || ""),
              ownerid: String(selectedRecord.ownerid || ""),
              ownerid_entitytype: String(selectedRecord.ownerid_entitytype || ""),
              owneridname: String(selectedRecord.owneridname || ""),
              new_phonecallreason: String(selectedRecord.new_phonecallreason || ""),
              new_phonecalloutcome: String(selectedRecord.new_phonecalloutcome || ""),
              strava_note: String(selectedRecord.strava_note || ""),
              strava_notecategory: String(selectedRecord.strava_notecategory || ""),
              strava_publishnote: String(selectedRecord.strava_note || ""),
              description: String(selectedRecord.description || ""),
              softphon_queuename: String(selectedRecord.softphon_queuename || ""),
              softphon_queuetimeseconds: String(selectedRecord.softphon_queuetimeseconds || ""),
              softphon_genesyscloudwrapupcode: String(selectedRecord.softphon_genesyscloudwrapupcode || ""),
              softphon_ivrtimeseconds: String(selectedRecord.softphon_ivrtimeseconds || ""),
              actualdurationminutes: String(selectedRecord.actualdurationminutes || ""),
              softphon_durationseconds: String(selectedRecord.softphon_durationseconds || ""),
              softphon_dispositiondurationseconds: String(selectedRecord.softphon_dispositiondurationseconds || ""),
              softphon_interactionurl: String(selectedRecord.softphon_interactionurl || ""),
            }}
            onClose={this.closeForm}
          />
        )}
        {isAppointmentFormOpen && selectedRecord && (
          <AppointmentForm
            data={{
              id: selectedRecord.id,
              subject: selectedRecord.subject,
              regarding: selectedRecord.regardingobjectidname,
              regardingobjectid_entitytype: String(selectedRecord.regardingobjectid_entitytype || ""),
              regardingobjectid: String(selectedRecord.regardingobjectid || ""),
              createdon: selectedRecord.createdon,
              ownerid: String(selectedRecord.ownerid || ""),
              ownerid_entitytype: String(selectedRecord.ownerid_entitytype || ""),
              owneridname: String(selectedRecord.owneridname || ""),
              status: selectedRecord.status,
              priority: selectedRecord.priority,
              scheduledend: selectedRecord.scheduledend,
              scheduledstart: String(selectedRecord.scheduledstart),
              activitytypecode: selectedRecord.activitytypecode as "Appointment",
              new_appttype: String(selectedRecord.new_appttype || ""),
              new_apptmethod: String(selectedRecord.new_apptmethod || "N/A"),
              strava_business: String(selectedRecord.strava_business || ""),
              strava_businessdirect: String(selectedRecord.strava_businessdirect || ""),
              strava_homenumber: String(selectedRecord.strava_homenumber || ""),
              strava_mobile: String(selectedRecord.strava_mobile || ""),
              strava_note: String(selectedRecord.strava_note || ""),
              strava_notecategory: String(selectedRecord.strava_notecategory || ""),
              strava_otherphone: String(selectedRecord.strava_otherphone || ""),
              strava_publishnote: Boolean(selectedRecord.strava_publishnote),
              new_appointmentoutcome: String(selectedRecord.new_appointmentoutcome || ""),
              scheduleddurationminutes: String(selectedRecord.scheduleddurationminutes || ""),
              location: String(selectedRecord.location || ""),
              isalldayevent: Boolean(selectedRecord.isalldayevent),
              requiredattendees: Array.isArray(selectedRecord.requiredattendees) ? selectedRecord.requiredattendees as IPartyList[] : [],
              requiredattendees_entitytype: String(selectedRecord.requiredattendees_entitytype || ""),
              requiredattendeesname: String(selectedRecord.requiredattendeesname || ""),
              optionalattendees: Array.isArray(selectedRecord.optionalattendees) ? selectedRecord.optionalattendees as IPartyList[] : [],
              optionalattendees_entitytype: String(selectedRecord.optionalattendees_entitytype || ""),
              optionalattendeesname: String(selectedRecord.optionalattendeesname || ""),
              description: String(selectedRecord.description || ""),
              isonlinemeeting: String(selectedRecord.isonlinemeeting),
              attachments: Array.isArray(selectedRecord.attachments) ? selectedRecord.attachments as IAttachment[] : [],
            }}
            onClose={this.closeForm}
          />
        )}
      </div>
    );
  }
}
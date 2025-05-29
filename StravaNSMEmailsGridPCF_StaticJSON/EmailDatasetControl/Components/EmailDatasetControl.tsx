import { IInputs, IOutputs } from "../generated/ManifestTypes";
import {
  DetailsList, DetailsListLayoutMode, Selection, IColumn, mergeStyleSets, TextField, IButtonStyles,
  Dropdown, IDropdownOption, PrimaryButton, DefaultButton, CheckboxVisibility, Link,
} from "@fluentui/react";
import * as React from "react";
import CustomEmailForm, { EmailFormProps } from "./Forms/CustomEmailForm";
import { emailData, fetchAllEmailData } from "./Backend/EmailData";


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
  selectedIds: string[];
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
  { key: "Email", text: "Email" },
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
      selectedIds: [],
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
      onSelectionChanged: this.onSelectionChanged,
      selectionMode: 2, // multiple
    });
  }

  onSelectionChanged = () => {
    const selectedItems = this._selection.getSelection() as IRecord[];
    this.setState({
      selectedItems,
      selectedIds: selectedItems.map((item) => item.id),
    });
  };

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
        key: "from",
        name: "From",
        fieldName: "from",
        minWidth: 120,
        maxWidth: 200,
        isResizable: true,
        headerClassName: classNames.headerWrapper,
        isSorted: false,
        isSortedDescending: false,
        onColumnClick: this.onColumnClick,
        onRender: (item: IRecord) =>
          Array.isArray(item.from) && item.from.length > 0
            ? (
              <>
                {(item.from as IPartyList[])
                  .map((party, idx) =>
                    party && party.partyid && party.entitytype ? (
                      <Link
                        key={party.partyid}
                        className={classNames.subjectLink}
                        onClick={() =>
                          this.openCrmForm(String(party.entitytype), String(party.partyid))
                        }
                        disabled={!party.partyid || !party.entitytype}
                      >
                        {party.name || "N/A"}
                      </Link>
                    ) : (
                      <span key={idx}>N/A</span>
                    )
                  )
                  .reduce<React.ReactNode[]>((acc, curr, idx) => {
                    if (idx === 0) return [curr];
                    return [...acc, "; ", curr];
                  }, [])
                }
              </>
            )
            : "N/A",
      },
      {
        key: "to",
        name: "To",
        fieldName: "to",
        minWidth: 120,
        maxWidth: 200,
        isResizable: true,
        headerClassName: classNames.headerWrapper,
        isSorted: false,
        isSortedDescending: false,
        onColumnClick: this.onColumnClick,
        onRender: (item: IRecord) =>
          Array.isArray(item.to) && item.to.length > 0
            ? (
              <>
                {(item.to as IPartyList[])
                  .map((party, idx) =>
                    party && party.partyid && party.entitytype ? (
                      <Link
                        key={party.partyid}
                        className={classNames.subjectLink}
                        onClick={() =>
                          this.openCrmForm(String(party.entitytype), String(party.partyid))
                        }
                        disabled={!party.partyid || !party.entitytype}
                      >
                        {party.name || "N/A"}
                      </Link>
                    ) : (
                      <span key={idx}>N/A</span>
                    )
                  )
                  .reduce<React.ReactNode[]>((acc, curr, idx) => {
                    if (idx === 0) return [curr];
                    return [...acc, "; ", curr];
                  }, [])
                }
              </>
            )
            : "N/A",
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

  private getEntitySetName(entityName: string): string {
    if (!entityName) return "";
    return entityName.endsWith("y")
      ? `${entityName.slice(0, -1)}ies`
      : `${entityName}s`;
  }

  private async restoreRecords(): Promise<void> {

    const { selectedItems } = this.state;
    const { context } = this.props;

    if (selectedItems.length === 0) {
      this.setState({ errorMessage: "No records selected for restoration." });
      return;
    }

    this.setState({ isLoading: true, errorMessage: null });

    // Array to store errors for each record
    const errors: { recordId: string; subject: string; errorMessage: string }[] = [];
    let successfulRestorations = 0;

    try {
      for (const record of selectedItems) {
        try {
          let entityName: string;
          let entityData: Record<string, unknown> = {
            activityid: record.activityid, // Use 'id' from IRecord
            subject: record.subject,
          };

          switch (record.activitytypecode) {
            case "Email": {
              entityName = "email";

              const activityParties: object[] = [];

              if (record.from && Array.isArray(record.from)) {
                (record.from as IPartyList[]).forEach((sender) => {
                  if (sender.partyid && sender.entitytype) {
                    const toSender = {
                      [`partyid_${sender.entitytype}@odata.bind`]: `/${this.getEntitySetName(String(sender.entitytype))}(${sender.partyid})`,
                      participationtypemask: 1,
                    };
                    activityParties.push(toSender);
                  }
                });
              }
              if (record.to && Array.isArray(record.to)) {
                (record.to as IPartyList[]).forEach((recipient) => {
                  if (recipient.partyid && recipient.entitytype) {
                    const toRecipient = {
                      [`partyid_${recipient.entitytype}@odata.bind`]: `/${this.getEntitySetName(String(recipient.entitytype))}(${recipient.partyid})`,
                      participationtypemask: 2,
                    };
                    activityParties.push(toRecipient);
                  }
                });
              }
              if (record.cc && Array.isArray(record.cc)) {
                (record.cc as IPartyList[]).forEach((recipient) => {
                  if (recipient.partyid && recipient.entitytype) {
                    const ccRecipient = {
                      [`partyid_${recipient.entitytype}@odata.bind`]: `/${this.getEntitySetName(String(recipient.entitytype))}(${recipient.partyid})`,
                      participationtypemask: 3,
                    };
                    activityParties.push(ccRecipient);
                  }
                });
              }
              if (record.bcc && Array.isArray(record.bcc)) {
                (record.bcc as IPartyList[]).forEach((recipient) => {
                  if (recipient.partyid && recipient.entitytype) {
                    const bccRecipient = {
                      [`partyid_${recipient.entitytype}@odata.bind`]: `/${this.getEntitySetName(String(recipient.entitytype))}(${recipient.partyid})`,
                      participationtypemask: 4,
                    };
                    activityParties.push(bccRecipient);
                  }
                });
              }

              entityData = {
                ...entityData,
                description: record.description,
                directioncode: true,
                statecode:record.statecode,
                statuscode: record.statuscode,
                prioritycode: record.prioritycode,
                ...(record.regardingobjectid && record.regardingobjectid_entitytype && {
                  [`regardingobjectid_${record.regardingobjectid_entitytype}_email@odata.bind`]: `/${this.getEntitySetName(String(record.regardingobjectid_entitytype))}(${record.regardingobjectid})`,
                }),
                ...(record.ownerid && record.ownerid_entitytype && {
                  ["ownerid_email@odata.bind"]: `/${this.getEntitySetName(String(record.ownerid_entitytype))}(${record.ownerid})`,
                }),
                ...(activityParties.length > 0 && { email_activity_parties: activityParties }),
              };
              break;
            }
            default:
              throw new Error(`Unsupported activity type: ${record.activitytypecode}`);
          }

          // Remove undefined or null values
          Object.keys(entityData).forEach((key) => {
            if (entityData[key] === undefined || entityData[key] === null) {
              delete entityData[key];
            }
          });

          // Create record using Web API
          const createResult = await context.webAPI.createRecord(entityName, entityData);
          const createdActivityId = createResult.id;

          // Handle attachments
          if (record.attachments && Array.isArray(record.attachments) && record.attachments.length > 0) {
            for (const attachment of record.attachments) {
              if (
                typeof attachment === "object" &&
                attachment !== null &&
                "body" in attachment &&
                "filename" in attachment &&
                "filesize" in attachment &&
                "mimetype" in attachment
              ) {
                await this.createAttachment(context, createdActivityId, entityName, {
                  body: (attachment as IAttachment).body,
                  filename: (attachment as IAttachment).filename,
                  filesize: (attachment as IAttachment).filesize,
                  mimetype: (attachment as IAttachment).mimetype,
                });
              }
            }
          }

          successfulRestorations++;
        } catch (recordError) {
          // Store error for this specific record
          const errorMessage = recordError instanceof Error ? recordError.message : "Unknown error";
          errors.push({
            recordId: record.id,
            subject: record.subject || "Unnamed Record",
            errorMessage,
          });
          console.error(`Error restoring record ${record.id}:`, recordError);
        }
      }

      // Refresh the view after processing all records
      await this.updateView();

      // Construct final message
      let finalMessage = "";
      if (successfulRestorations > 0) {
        finalMessage += `${successfulRestorations} record${successfulRestorations > 1 ? "s" : ""} restored successfully. `;
      }
      if (errors.length > 0) {
        finalMessage += `Failed to restore ${errors.length} record${errors.length > 1 ? "s" : ""}:\n`;
        errors.forEach((error, index) => {
          finalMessage += `${index + 1}. Record "${error.subject}" (ID: ${error.recordId}): ${error.errorMessage}\n`;
        });
      } else {
        finalMessage = "All records restored successfully.";
      }

      this.setState({
        isLoading: false,
        errorMessage: finalMessage,
        selectedItems: [],
        selectedIds: [],
      });
      this._selection.setAllSelected(false);

      // Clear message after 10 seconds
      setTimeout(() => {
        this.setState({ errorMessage: null });
      }, 10000);
    } catch (generalError) {
      console.error("Unexpected error during restoration:", generalError);
      this.setState({
        isLoading: false,
        errorMessage: `Unexpected error during restoration: ${generalError instanceof Error ? generalError.message : "Unknown error"}`,
      });
    }
  }

  private async createAttachment(
    context: ComponentFramework.Context<IInputs>,
    activityId: string,
    activityType: string,
    attachment: {
      body: string;
      filename: string;
      filesize: number;
      mimetype: string;
    }
  ): Promise<void> {
    const attachmentData = {
      ["objectid_activitypointer@odata.bind"]: `/activitypointers(${activityId})`,
      objecttypecode: activityType,
      body: attachment.body,
      subject: attachment.filename,
      filename: attachment.filename,
      mimetype: attachment.mimetype,
      filesize: attachment.filesize,
    };

    await context.webAPI.createRecord("activitymimeattachment", attachmentData);
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

  async componentDidMount() {
    this.setState({ isLoading: true });
    try {
      const allEmails = await fetchAllEmailData(this.context as ComponentFramework.Context<IInputs>);
      this.setState({
        items: allEmails.map(email => ({
          ...email,
          activitytypecode: "Email" as "Email"
        })),
        isLoading: false
      }, () => {
        this.updateView();
      });
    } catch (error) {
      this.setState({
        items: [],
        filteredItems: [],
        isLoading: false,
        errorMessage: "Failed to load email data.",
      });
    }
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

    try {
      const items: IRecord[] = [
        ...this.state.items.map((email) => ({
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
      ];

      const filteredItems = this.applyFilters(items);

      this.setState({
        items,
        filteredItems,
        isLoading: false,
        errorMessage: items.length === 0 ? "No records found for this ID." : null,
      });
    } catch (error) {
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
      let aValue = a[sortColumn];
      let bValue = b[sortColumn];

      // Handle array values for "from" and "to"
      if ((sortColumn === "from" || sortColumn === "to") && Array.isArray(aValue) && Array.isArray(bValue)) {
        // Join names for comparison
        aValue = (aValue as IPartyList[]).map((p: IPartyList) => p?.name || "").join("; ");
        bValue = (bValue as IPartyList[]).map((p: IPartyList) => p?.name || "").join("; ");
      }

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

      const comparison = (aValue ?? "").toString().localeCompare((bValue ?? "").toString());
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
    this.setState({ currentPage: page }, () => {
      this.updateSelectionForPage();
    });
  };

  updateSelectionForPage = () => {
    const { filteredItems, currentPage, pageSize, selectedIds } = this.state;
    const startIndex = (currentPage - 1) * pageSize;
    const endIndex = startIndex + pageSize;
    const paginatedItems = filteredItems.slice(startIndex, endIndex);

    // Deselect all, then select only those on this page that are in selectedIds
    this._selection.setAllSelected(false);
    paginatedItems.forEach((item, idx) => {
      if (selectedIds.includes(item.id)) {
        this._selection.setIndexSelected(idx, true, false);
      }
    });
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
            {/* <label
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
            /> */}
          </div>
          <div style={{ display: "flex", gap: "10px", alignItems: "center" }}>
            <PrimaryButton
              text="Restore"
              iconProps={{ iconName: "Refresh" }}
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
              from: Array.isArray(selectedRecord.from) ? selectedRecord.from as IPartyList[] : [],
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
      </div>
    );
  }
}
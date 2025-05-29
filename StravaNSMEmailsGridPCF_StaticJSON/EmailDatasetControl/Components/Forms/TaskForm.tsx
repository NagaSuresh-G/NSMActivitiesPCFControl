import {
    CommandBar,
    Label,
    Modal,
    Stack,
    TextField,
    IconButton,
} from "@fluentui/react";
import * as React from "react";

export interface ITaskRecord {
    id: string;
    subject: string;
    strava_url: string;
    regarding: string;
    regardingobjectid_entitytype: string;
    regardingobjectid: string;
    strava_contact: string;
    strava_contactname?: string;
    strava_contact_entitytype: string;
    strava_business?: string;
    strava_businessdirect?: string;
    strava_otherphone?: string;
    strava_homenumber?: string;
    strava_mobile?: string;
    scheduledend?: string;
    priority?: string;
    ownerid?: string;
    ownerid_entitytype: string;
    owneridname: string;
    strava_connectedcallstatus?: string;
    strava_createoutlookcalendarreminder: string;
    strava_notecategory?: string;
    actualdurationminutes?: string;
    strava_note?: string;
    description?: string;
    createdon: string;
    status: string;
    activitytypecode: string;
    strava_publishnote: boolean;
    new_taskpriority?: string;
    new_teamassociation?: string;
    new_quoteboundindicator?: string;
}

export interface TaskFormProps {
    data: ITaskRecord;
    onClose: () => void;
}

const openCrmForm = (entityName: string, id: string) => {
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

const taskForm: React.FC<TaskFormProps> = ({ data, onClose }) => {
    return (
        <Modal
            titleAriaId={`Task-${data.id}`}
            isOpen={true}
            onDismiss={onClose}
            isBlocking={false}
            styles={{
                main: {
                    width: "95vw",
                    maxWidth: "1500px",
                    height: "95vh",
                    maxHeight: "950px",
                    padding: "0",
                    display: "flex",
                    overflow: "hidden",
                },
                scrollableContent: {
                    padding: "0",
                    overflowY: "hidden",
                    height: "100%",
                },
            }}
        >
            <Stack tokens={{ childrenGap: 0 }} styles={{ root: { height: "100%", overflowY: "auto" } }}>
                {/* Header */}
                <div
                    style={{
                        padding: "16px 24px",
                        backgroundColor: "#f3f2f1",
                        borderBottom: "1px solid #e1e1e1",
                        display: "flex",
                        justifyContent: "space-between",
                        alignItems: "center",
                    }}
                >
                    <h2 style={{ margin: 0, fontSize: "20px" }}>
                        {data.subject}
                    </h2>
                    <Stack
                        horizontal
                        tokens={{ childrenGap: 10 }}
                        styles={{
                            root: {
                                alignItems: "center",
                                flexWrap: "nowrap",
                            },
                        }}
                    >
                        <TextField
                            label="Priority"
                            readOnly
                            value={data.priority || "Not specified"}
                            styles={{ root: { width: "100px" } }}
                        />
                        <TextField label="Status" readOnly value={data.status} styles={{ root: { width: "100px" } }} />
                        <div style={{ display: "flex", flexDirection: "column", width: "130px" }}>
                            <Label>Owner</Label>
                            <div
                                style={{
                                    padding: "6px 12px",
                                    minHeight: "20px",
                                    display: "flex",
                                    alignItems: "center",
                                    border: "1px solid black",
                                    borderRadius: "2px",
                                    background: "#FFFFFF",
                                    overflow: "hidden",
                                    whiteSpace: "nowrap",
                                    textOverflow: "ellipsis",
                                }}
                                title={data.owneridname || "Not specified"} // This shows full text on hover
                            >
                                <a
                                    href="#"
                                    onClick={(e) => {
                                        e.preventDefault();
                                        openCrmForm(data.ownerid_entitytype, data.ownerid as string);
                                    }}
                                    style={{
                                        textDecoration: "underline",
                                        color: "#0078d4",
                                        overflow: "hidden",
                                        whiteSpace: "nowrap",
                                        textOverflow: "ellipsis",
                                        display: "inline-block",
                                        width: "100%",
                                    }}
                                    title={data.owneridname || "Not specified"}
                                >
                                    {data.owneridname || "Not specified"}
                                </a>
                            </div>
                        </div>
                        <IconButton
                            iconProps={{ iconName: "Cancel" }}
                            ariaLabel="Close modal"
                            onClick={onClose}
                            styles={{
                                root: {
                                    color: "#666",
                                    height: "32px",
                                    width: "32px",
                                },
                                rootHovered: {
                                    backgroundColor: "#e1e1e1",
                                    color: "#000",
                                },
                            }}
                        />
                    </Stack>
                </div>

                {/* Main content: Form fields and Placeholder for Attachments/Related Data */}
                <Stack horizontal tokens={{ childrenGap: 20 }} styles={{ root: { padding: "16px 24px", flexWrap: "wrap" } }}>
                    {/* Left Column */}
                    <Stack tokens={{ childrenGap: 15 }} styles={{ root: { flex: 7, minWidth: "500px" } }}>
                        {/* Group 1: Subject, URL, Regarding */}
                        <Stack tokens={{ childrenGap: 10 }} styles={{ root: { border: "1px solid #e1dfdd", padding: "10px", borderRadius: "2px" } }}>
                            <TextField label="Subject" readOnly value={data.subject} />
                            <TextField label="URL" readOnly value={data.strava_url} />
                            {/* <TextField label="Regarding" readOnly value={data.regarding} /> */}
                            <div style={{ marginBottom: 12 }}>
                                <Label>Regarding</Label>
                                <div
                                    style={{
                                        padding: "6px 12px",
                                        height: "18px",
                                        display: "flex",
                                        alignItems: "center",
                                        border: "1px solid black",
                                        borderRadius: "2px",
                                        background: "##FFFFFF",
                                    }}
                                >
                                    <a
                                        href="#"
                                        onClick={(e) => {
                                            e.preventDefault();
                                            openCrmForm(data.regardingobjectid_entitytype, data.regardingobjectid);
                                        }}
                                        style={{ textDecoration: "underline", color: "#0078d4" }}
                                    >
                                        {data.regarding}
                                    </a>
                                </div>
                            </div>
                        </Stack>
                        {/* Group 2: Contact-related fields */}
                        <Stack tokens={{ childrenGap: 10 }} styles={{ root: { border: "1px solid #e1dfdd", padding: "10px", borderRadius: "2px" } }}>
                            <Stack horizontal tokens={{ childrenGap: 10 }}>
                                {/* <TextField label="Contact" readOnly value={data.strava_contact} styles={{ root: { flex: 1 } }} /> */}
                                <div style={{ marginBottom: 12, flex: 1 }}>
                                <Label>Contact</Label>
                                <div
                                    style={{
                                        padding: "6px 12px",
                                        height: "18px",
                                        display: "flex",
                                        alignItems: "center",
                                        border: "1px solid black",
                                        borderRadius: "2px",
                                        background: "##FFFFFF",
                                    }}
                                >
                                    <a
                                        href="#"
                                        onClick={(e) => {
                                            e.preventDefault();
                                            openCrmForm(data.strava_contact_entitytype, data.strava_contact);
                                        }}
                                        style={{ textDecoration: "underline", color: "#0078d4" }}
                                    >
                                        {data.strava_contactname}
                                    </a>
                                </div>
                            </div>
                                <TextField label="Business" readOnly value={data.strava_business || "---"} styles={{ root: { flex: 1 } }} />
                            </Stack>
                            <Stack horizontal tokens={{ childrenGap: 10 }}>
                                <TextField label="Home" readOnly value={data.strava_homenumber} styles={{ root: { flex: 1 } }} />
                                <TextField label="Business Direct" readOnly value={data.strava_businessdirect || "---"} styles={{ root: { flex: 1 } }} />
                            </Stack>
                            <Stack horizontal tokens={{ childrenGap: 10 }}>
                                <TextField label="Mobile" readOnly value={data.strava_mobile || "---"} styles={{ root: { flex: 1 } }} />
                                <TextField label="Other Phone" readOnly value={data.strava_otherphone || "---"} styles={{ root: { flex: 1 } }} />
                            </Stack>
                        </Stack>
                        <Stack tokens={{ childrenGap: 10 }} styles={{ root: { border: "1px solid #e1dfdd", padding: "10px", borderRadius: "2px" } }}>
                            <Label>Description</Label>
                            <div style={{ border: "1px solid #e1dfdd", padding: "10px", height: "150px", overflowY: "auto", whiteSpace: "pre-wrap", borderRadius: "2px" }} dangerouslySetInnerHTML={{ __html: data.description || "<i>No description provided</i>" }} />
                            <TextField label="Task Priority" readOnly value={data.new_taskpriority} />
                            <div style={{ marginBottom: 12 }}>
                                <Label>Owner</Label>
                                <div
                                    style={{
                                        padding: "6px 12px",
                                        height: "18px",
                                        display: "flex",
                                        alignItems: "center",
                                        border: "1px solid black",
                                        borderRadius: "2px",
                                        background: "##FFFFFF",
                                    }}
                                >
                                    <a
                                        href="#"
                                        onClick={(e) => {
                                            e.preventDefault();
                                            openCrmForm(data.ownerid_entitytype, data.ownerid as string);
                                        }}
                                        style={{ textDecoration: "underline", color: "#0078d4" }}
                                    >
                                        {data.owneridname}
                                    </a>
                                </div>
                            </div>
                            <TextField label="Team Association" readOnly value={data.new_teamassociation} />
                            <TextField label="Duration" readOnly value={data.actualdurationminutes} />
                            <TextField label="Quote Bound Indicator" readOnly value={data.new_quoteboundindicator} />
                        </Stack>
                    </Stack>
                    {/* Right Column */}
                    <Stack tokens={{ childrenGap: 15 }} styles={{ root: { flex: 3, minWidth: "400px" } }}>
                        <Stack tokens={{ childrenGap: 10 }} styles={{ root: { border: "1px solid #e1dfdd", padding: "10px", borderRadius: "2px" } }}>
                        <TextField label="Task Category" readOnly value={data.strava_notecategory || "---"} />
                        <TextField label="Due Date" readOnly value={data.scheduledend || "---"} />
                        <TextField label="Connected Call Status" readOnly value={data.strava_connectedcallstatus || "---"} />
                        <TextField label="Create Outlook Calendar Reminder" readOnly value={data.strava_createoutlookcalendarreminder || "---"} />
                        <TextField label="Note Category" readOnly value={data.strava_notecategory || "---"} />
                        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                            <Label>Publish Note? (Y/N)</Label>
                            <input type="checkbox" disabled checked={data.strava_publishnote} />
                        </Stack>
                        <TextField label="Note" readOnly value={data.strava_note || "---"} multiline /></Stack>
                    </Stack>
                </Stack>

            </Stack>
        </Modal>
    );
};

export default taskForm;
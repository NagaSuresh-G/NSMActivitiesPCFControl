import {
    Label,
    Modal,
    Stack,
    TextField,
    IconButton,
    DocumentCard,
    DocumentCardDetails,
    Icon,
} from "@fluentui/react";
import * as React from "react";

export interface IRecord {
    id: string;
    subject: string;
    regardingobjectid: string;
    regardingobjectid_entitytype: string;
    regarding: string;
    priority: string;
    status: string;
    description?: string;
    ownerid?: string;
    ownerid_entitytype: string;
    owneridname: string;
    requiredattendees: IPartyList[];
    requiredattendees_entitytype: string;
    requiredattendeesname?: string;
    optionalattendees: IPartyList[];
    optionalattendees_entitytype: string;
    optionalattendeesname?: string;
    new_appttype?: string;
    strava_homenumber?: string;
    strava_mobile?: string;
    strava_business?: string;
    strava_businessdirect?: string;
    strava_otherphone?: string;
    new_apptmethod?: string;
    new_appointmentoutcome?: string;
    isonlinemeeting?: string;
    createdon: string;
    activitytypecode: string;
    scheduledend: string;
    scheduledstart: string;
    location?: string;
    strava_notecategory?: string;
    strava_publishnote?: boolean;
    strava_note?: string;
    scheduleddurationminutes?: string;
    isalldayevent?: boolean;
    attachments?: IAttachment[];
    [key: string]: string | number | boolean | null | undefined | IAttachment[] | IPartyList[];
}

export interface IAttachment {
    attachmentId: string;
    body: string;
    filename: string;
    filesize: number;
    mimetype: string;
}

export interface IPartyList {
    partyid: string;
    name: string;
    entitytype: string;
}

export interface AppointmentFormProps {
    data: IRecord;
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

const AppointmentForm: React.FC<AppointmentFormProps> = ({ data, onClose }) => {
    const handleAttachmentClick = (attachment: IAttachment) => {
        const link = document.createElement("a");
        link.href = `data:${attachment.mimetype};base64,${attachment.body}`;
        link.download = attachment.filename;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    };

    return (
        <Modal
            titleAriaId={`Appointment-${data.id}`}
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
                                alignItems: "laubt",
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
                                title={data.owneridname || "Not specified"}
                            >
                                <a
                                    href="#"
                                    onClick={(e: React.MouseEvent<HTMLAnchorElement>) => {
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

                {/* Main content: Form fields and Attachments */}
                <Stack
                    horizontal
                    tokens={{ childrenGap: 20 }}
                    styles={{ root: { padding: "16px 24px", flexWrap: "wrap" } }}
                >
                    {/* Left section: Form fields */}
                    <Stack
                        tokens={{ childrenGap: 15 }}
                        styles={{ root: { flex: 1, minWidth: "300px" } }}
                    >
                        <Stack tokens={{ childrenGap: 10 }} styles={{ root: { border: "1px solid #e1dfdd", padding: "10px", borderRadius: "2px" } }}>
                            <div style={{ marginBottom: 12 }}>
                                <Label>Required</Label>
                                <div
                                    style={{
                                        padding: "6px 12px",
                                        height: "18px",
                                        display: "flex",
                                        alignItems: "center",
                                        border: "1px solid black",
                                        borderRadius: "2px",
                                        background: "#FFFFFF",
                                    }}
                                >
                                    {data.requiredattendees && data.requiredattendees.length > 0 ? (
                                        data.requiredattendees.map((required, index) => (
                                            <React.Fragment key={required.partyid}>
                                                <a
                                                    href="#"
                                                    onClick={(e) => {
                                                        e.preventDefault();
                                                        openCrmForm(required.entitytype, required.partyid);
                                                    }}
                                                    style={{
                                                        textDecoration: "underline",
                                                        color: "#0078d4",
                                                        marginRight: "4px",
                                                    }}
                                                    title={required.name}
                                                >
                                                    {required.name}
                                                </a>
                                                {index < data.requiredattendees.length - 1 && (
                                                    <span style={{ marginRight: "4px" }}>;</span>
                                                )}
                                            </React.Fragment>
                                        ))
                                    ) : (
                                        <span style={{ color: "#666" }}>No recipients</span>
                                    )}
                                </div>
                            </div>
                            <div style={{ marginBottom: 12 }}>
                                <Label>Optional</Label>
                                <div
                                    style={{
                                        padding: "6px 12px",
                                        height: "18px",
                                        display: "flex",
                                        alignItems: "center",
                                        border: "1px solid black",
                                        borderRadius: "2px",
                                        background: "#FFFFFF",
                                    }}
                                >
                                    {data.optionalattendees && data.optionalattendees.length > 0 ? (
                                        data.optionalattendees.map((optional, index) => (
                                            <React.Fragment key={optional.partyid}>
                                                <a
                                                    href="#"
                                                    onClick={(e) => {
                                                        e.preventDefault();
                                                        openCrmForm(optional.entitytype, optional.partyid);
                                                    }}
                                                    style={{
                                                        textDecoration: "underline",
                                                        color: "#0078d4",
                                                        marginRight: "4px",
                                                    }}
                                                    title={optional.name}
                                                >
                                                    {optional.name}
                                                </a>
                                                {index < data.optionalattendees.length - 1 && (
                                                    <span style={{ marginRight: "4px" }}>;</span>
                                                )}
                                            </React.Fragment>
                                        ))
                                    ) : (
                                        <span style={{ color: "#666" }}>No recipients</span>
                                    )}
                                </div>
                            </div>
                            <TextField label="Subject" readOnly value={data.subject || "Not specified"} />
                            <TextField label="Location" readOnly value={data.location || "Not specified"} />
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
                                        background: "#FFFFFF",
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
                                        {data.regarding || "Not specified"}
                                    </a>
                                </div>
                            </div>
                            <Stack horizontal tokens={{ childrenGap: 10 }}>
                                <TextField label="Appt Type" readOnly value={data.new_appttype} styles={{ root: { flex: 1 } }} />
                                <TextField label="Home" readOnly value={data.strava_homenumber || "---"} styles={{ root: { flex: 1 } }} />
                            </Stack>
                            <Stack horizontal tokens={{ childrenGap: 10 }}>
                                <TextField label="Appt Method" readOnly value={data.new_apptmethod} styles={{ root: { flex: 1 } }} />
                                <TextField label="Mobile" readOnly value={data.strava_mobile || "---"} styles={{ root: { flex: 1 } }} />
                            </Stack>
                            <Stack horizontal tokens={{ childrenGap: 10 }}>
                                <TextField label="Appointment Outcome" readOnly value={data.new_appointmentoutcome || "---"} styles={{ root: { flex: 1 } }} />
                                <TextField label="Business" readOnly value={data.strava_business || "---"} styles={{ root: { flex: 1 } }} />
                            </Stack>
                            <Stack horizontal tokens={{ childrenGap: 10 }}>
                                <TextField label="Business Direct" readOnly value={data.strava_businessdirect || "---"} styles={{ root: { flex: 1 } }} />
                                <TextField label="OtherPhone" readOnly value={data.strava_otherphone || "---"} styles={{ root: { flex: 1 } }} />
                            </Stack>
                            <Stack horizontal tokens={{ childrenGap: 10 }}>
                                <TextField label="Created On" readOnly value={data.createdon || "---"} styles={{ root: { flex: 1 } }} />
                                <TextField label="Teams meeting" readOnly value={data.isonlinemeeting || "---"} styles={{ root: { flex: 1 } }} />
                            </Stack>
                        </Stack>

                        <Stack tokens={{ childrenGap: 10 }} styles={{ root: { border: "1px solid #e1dfdd", padding: "10px", borderRadius: "2px" } }}>
                            <Label>Description</Label>
                            <div
                                style={{
                                    border: "1px solid #c8c6c4",
                                    padding: "10px",
                                    height: "200px",
                                    overflowY: "auto",
                                    whiteSpace: "pre-wrap",
                                    borderRadius: "2px",
                                }}
                                dangerouslySetInnerHTML={{
                                    __html: data.description || "<i>No description provided</i>",
                                }}
                            />
                        </Stack>

                        <Stack tokens={{ childrenGap: 10 }} styles={{ root: { border: "1px solid #e1dfdd", padding: "10px", borderRadius: "2px" } }}>
                            <TextField label="Note Category" readOnly value={data.strava_notecategory || "---"} />
                            <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                                <Label>Publish Note? (Y/N)</Label>
                                <input type="checkbox" disabled checked={data.strava_publishnote} />
                            </Stack>
                            <TextField label="Note" readOnly value={data.strava_note || "---"} multiline />
                        </Stack>

                        <Stack tokens={{ childrenGap: 10 }} styles={{ root: { border: "1px solid #e1dfdd", padding: "10px", borderRadius: "2px" } }}>
                            <TextField label="Start Time" readOnly value={data.scheduledstart || "---"} />
                            <TextField label="End Time" readOnly value={data.scheduledend || "---"} />
                            <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                                <Label>All Day Event</Label>
                                <input type="checkbox" disabled checked={data.isalldayevent} />
                            </Stack>
                            <TextField label="Duration" readOnly value={data.scheduleddurationminutes || "---"} />
                        </Stack>
                    </Stack>

                    {/* Right section: Attachments */}
                    <Stack
                        styles={{
                            root: {
                                width: "300px",
                                height: "400px",
                                border: "1px solid #e1e1e1",
                                padding: "16px",
                                borderRadius: "4px",
                                overflowY: "auto",
                            },
                        }}
                    >
                        <Stack horizontal horizontalAlign="space-between" styles={{ root: { marginBottom: "8px" } }}>
                            <Label>Attachments ({data.attachments ? data.attachments.length : 0})</Label>
                        </Stack>
                        {data.attachments && data.attachments.length > 0 ? (
                            data.attachments.map((attachment: IAttachment) => (
                                <DocumentCard
                                    key={attachment.attachmentId}
                                    aria-label={`Attachment: ${attachment.filename}`}
                                    styles={{
                                        root: {
                                            width: "100%",
                                            marginBottom: "8px",
                                            paddingBottom: "0px",
                                        },
                                    }}
                                >
                                    <DocumentCardDetails>
                                        <Stack horizontal tokens={{ childrenGap: 8 }} styles={{ root: { alignItems: "center", margin: "0 8px" } }}>
                                            <Icon iconName="Attach" styles={{ root: { fontSize: "16px", color: "#0078d4" } }} />
                                            <a
                                                href={`data:${attachment.mimetype};base64,${attachment.body}`}
                                                download={attachment.filename}
                                                style={{
                                                    textDecoration: "underline",
                                                    color: "#0078d4",
                                                    cursor: "pointer",
                                                    fontSize: "14px",
                                                }}
                                                onClick={(e) => {
                                                    e.preventDefault();
                                                    handleAttachmentClick(attachment);
                                                }}
                                            >
                                                {attachment.filename}
                                            </a>
                                            <span style={{ fontSize: "12px", color: "#666" }}>
                                                ({(attachment.filesize / 1024).toFixed(2)} KB)
                                            </span>
                                        </Stack>
                                    </DocumentCardDetails>
                                </DocumentCard>
                            ))
                        ) : (
                            <div style={{ padding: "8px 0", color: "#666" }}>No attachments included.</div>
                        )}
                    </Stack>
                </Stack>
            </Stack>
        </Modal>
    );
};

export default AppointmentForm;
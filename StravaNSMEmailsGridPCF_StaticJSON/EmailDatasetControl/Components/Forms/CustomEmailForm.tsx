import {
    CommandBar,
    DocumentCard,
    DocumentCardActivity,
    DocumentCardDetails,
    DocumentCardTitle,
    Label,
    Modal,
    Pivot,
    PivotItem,
    Stack,
    TextField,
    ICommandBarItemProps,
    IconButton,
    Icon,
} from "@fluentui/react";
import * as React from "react";
import { IInputs } from "../../generated/ManifestTypes"; // adjust path if needed

// Interfaces remain the same
export interface IAttachment {
    attachmentId: string;
    body: string; // Base64 encoded attachment content
    filename: string;
    filesize: number;
    mimetype: string;
}

export interface IPartyList {
    partyid: string;
    name: string;
    entitytype: string;
}

export interface IRecord {
    id: string;
    subject: string;
    createdon: string;
    directioncode: string;
    from: IPartyList[];
    sendermailboxidname: string;
    to: IPartyList[];
    torecipients: string;
    regarding: string;
    regardingobjectid_entitytype: string;
    regardingobjectid: string;
    priority: string;
    status: string;
    description?: string;
    ownerid?: string;
    ownerid_entitytype: string;
    owneridname: string;
    cc: IPartyList[];
    bcc: IPartyList[];
    senton: string;
    attachments?: IAttachment[];
    [key: string]: string | number | boolean | null | undefined | IAttachment[] | IPartyList[];
}

export interface EmailFormProps {
    data: IRecord;
    onClose: () => void;
    context: ComponentFramework.Context<IInputs>;
}

const CustomEmailForm: React.FC<EmailFormProps> = ({ data, onClose, context }) => {
    const handleAttachmentClick = (attachment: IAttachment) => {
        const link = document.createElement("a");
        link.href = `data:${attachment.mimetype};base64,${attachment.body}`;
        link.download = attachment.filename;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    };

    const openCrmForm = (entityName: string, id: string) => {
        const orgUrl = "https://nsminc-sandbox.crm.dynamics.com";
        const formUrl = `${orgUrl}/main.aspx?etn=${encodeURIComponent(entityName)}&id=${encodeURIComponent(id)}&pagetype=entityrecord`;
        try {
            window.open(formUrl, "_blank");
            console.log(`Opened ${entityName} form in a new tab successfully.`);
        } catch (error) {
            const errMsg = error instanceof Error ? error.message : String(error);
            console.error(`Error opening ${entityName} form in a new tab: ${errMsg}`);
        }
    };

    return (
        <Modal
            titleAriaId={`email-${data.id}`}
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
                    <h2 style={{ margin: 0, fontSize: "20px" }}>{data.subject}</h2>
                    <Stack horizontal tokens={{ childrenGap: 10 }} styles={{ root: { alignItems: "center", flexWrap: "nowrap" } }}>
                        <TextField label="Priority" readOnly value={data.priority} styles={{ root: { width: "100px" } }} />
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

                {/* Main content: Form fields and Attachments */}
                <Stack horizontal tokens={{ childrenGap: 20 }} styles={{ root: { padding: "16px 24px", flexWrap: "wrap" } }}>
                    {/* Left section: Form fields */}
                    <Stack tokens={{ childrenGap: 15 }} styles={{ root: { flex: 1, minWidth: "300px",border: "1px solid #e1dfdd", padding: "10px", borderRadius: "2px" } }}>
                        <div style={{ marginBottom: 12 }}>
                            <Label>From</Label>
                            <div
                                style={{
                                    padding: "6px 12px",
                                    minHeight: "18px",
                                    display: "flex",
                                    alignItems: "center",
                                    border: "1px solid black",
                                    borderRadius: "2px",
                                    background: "#FFFFFF",
                                    flexWrap: "wrap",
                                }}
                            >
                                {data.from && data.from.length > 0 ? (
                                    data.from.map((sender, index) => (
                                        <React.Fragment key={sender.partyid}>
                                            {/* <a
                                                href="#"
                                                onClick={(e) => {
                                                    e.preventDefault();
                                                    openCrmForm(sender.entitytype, sender.partyid);
                                                }}
                                                style={{
                                                    textDecoration: "underline",
                                                    color: "#0078d4",
                                                    marginRight: "4px",
                                                }}
                                                title={sender.name}
                                            > */}
                                                {sender.name}
                                            {/* </a> */}
                                            {index < data.from.length - 1 && (
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
                            <Label>To</Label>
                            <div
                                style={{
                                    padding: "6px 12px",
                                    minHeight: "18px",
                                    display: "flex",
                                    alignItems: "center",
                                    border: "1px solid black",
                                    borderRadius: "2px",
                                    background: "#FFFFFF",
                                    flexWrap: "wrap",
                                }}
                            >
                                {data.to && data.to.length > 0 ? (
                                    data.to.map((recipient, index) => (
                                        <React.Fragment key={recipient.partyid}>
                                            {/* <a
                                                href="#"
                                                onClick={(e) => {
                                                    e.preventDefault();
                                                    openCrmForm(recipient.entitytype, recipient.partyid);
                                                }}
                                                style={{
                                                    textDecoration: "underline",
                                                    color: "#0078d4",
                                                    marginRight: "4px",
                                                }}
                                                title={recipient.name}
                                            > */}
                                                {recipient.name}
                                            {/* </a> */}
                                            {index < data.to.length - 1 && (
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
                            <Label>Cc</Label>
                            <div
                                style={{
                                    padding: "6px 12px",
                                    minHeight: "18px",
                                    display: "flex",
                                    alignItems: "center",
                                    border: "1px solid black",
                                    borderRadius: "2px",
                                    background: "#FFFFFF",
                                    flexWrap: "wrap",
                                }}
                            >
                                {data.cc && data.cc.length > 0 ? (
                                    data.cc.map((recipient, index) => (
                                        <React.Fragment key={recipient.partyid}>
                                            {/* <a
                                                href="#"
                                                onClick={(e) => {
                                                    e.preventDefault();
                                                    openCrmForm(recipient.entitytype, recipient.partyid);
                                                }}
                                                style={{
                                                    textDecoration: "underline",
                                                    color: "#0078d4",
                                                    marginRight: "4px",
                                                }}
                                                title={recipient.name}
                                            > */}
                                                {recipient.name}
                                            {/* </a> */}
                                            {index < data.cc.length - 1 && (
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
                            <Label>Bcc</Label>
                            <div
                                style={{
                                    padding: "6px 12px",
                                    minHeight: "18px",
                                    display: "flex",
                                    alignItems: "center",
                                    border: "1px solid black",
                                    borderRadius: "2px",
                                    background: "#FFFFFF",
                                    flexWrap: "wrap",
                                }}
                            >
                                {data.bcc && data.bcc.length > 0 ? (
                                    data.bcc.map((recipient, index) => (
                                        <React.Fragment key={recipient.partyid}>
                                            {/* <a
                                                href="#"
                                                onClick={(e) => {
                                                    e.preventDefault();
                                                    openCrmForm(recipient.entitytype, recipient.partyid);
                                                }}
                                                style={{
                                                    textDecoration: "underline",
                                                    color: "#0078d4",
                                                    marginRight: "4px",
                                                }}
                                                title={recipient.name}
                                            > */}
                                                {recipient.name}
                                            {/* </a> */}
                                            {index < data.bcc.length - 1 && (
                                                <span style={{ marginRight: "4px" }}>;</span>
                                            )}
                                        </React.Fragment>
                                    ))
                                ) : (
                                    <span style={{ color: "#666" }}>No recipients</span>
                                )}
                            </div>
                        </div>
                        <TextField label="Direction Code" readOnly value={data.directioncode} />
                        <TextField label="Subject" readOnly value={data.subject} />
                        <TextField label="Created On" readOnly value={data.createdon} />
                        <TextField label="Sent On" readOnly value={data.senton} />
                        <div style={{ marginBottom: 12 }}>
                            <Label>Regarding</Label>
                            <div
                                style={{
                                    padding: "6px 12px",
                                    height: "20px",
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
                                    {data.regarding}
                                </a>
                            </div>
                        </div>
                        <Label>Description</Label>
                        <div
                            style={{
                                border: "1px solid #c8c6c4",
                                padding: "10px",
                                height: "350px",
                                overflowY: "auto",
                                whiteSpace: "pre-wrap",
                                borderRadius: "2px",
                            }}
                            dangerouslySetInnerHTML={{
                                __html: data.description || "<i>No description provided</i>",
                            }}
                        />
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

export default CustomEmailForm;
import {
    Label,
    Modal,
    Stack,
    TextField,
    IconButton,
} from "@fluentui/react";
import * as React from "react";

export interface IRecord {
    id: string;
    subject: string;
    from: string;
    fromname: string;
    from_entitytype: string;
    to: IPartyList[];
    toname: string;
    to_entitytype: string;
    regarding: string;
    regardingobjectid: string;
    regardingobjectid_entitytype: string;
    priority: string;
    status: string;
    description?: string;
    ownerid: string;
    ownerid_entitytype: string;
    owneridname?: string;
    createdon: string;
    activitytypecode: string;
    scheduledend: string;
    phonenumber?: string;
    actualdurationminutes?: string;
    actualstart?: string;
    actualend?: string;
    directioncode: string;
    new_phonecallreason?: string;
    new_phonecalloutcome?: string;
    strava_note?: string;
    strava_notecategory?: string;
    strava_publishnote?: string;
    softphon_queuename?: string;
    softphon_queuetimeseconds?: string;
    softphon_genesyscloudwrapupcode?: string;
    softphon_ivrtimeseconds?: string;
    softphon_durationseconds?: string;
    softphon_dispositiondurationseconds?: string;
    softphon_interactionurl?: string;
    attachments?: IAttachment[];
    [key: string]: string | number | boolean | null | undefined | IAttachment[]| IPartyList[];
}

export interface IAttachment {
    attachmentId: string;
    body: string;
    filename: string;
    filesize: number;
    mimetype: string;
}
export interface IPartyList {
    partyid: string; name: string; entitytype: string;
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

export interface PhoneCallFormProps {
    data: IRecord;
    onClose: () => void;
}

const PhoneCallForm: React.FC<PhoneCallFormProps> = ({ data, onClose }) => {
    return (
        <Modal
            titleAriaId={`PhoneCall-${data.id}`}
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
                        <TextField
                            label="Status"
                            readOnly
                            value={data.status || "Not specified"}
                            styles={{ root: { width: "100px" } }}
                        />
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
                    <Stack tokens={{ childrenGap: 15 }} styles={{ root: { flex: 4, minWidth: "500px" } }}>
                        <Stack tokens={{ childrenGap: 10 }} styles={{ root: { border: "1px solid #e1dfdd", padding: "10px", borderRadius: "2px" } }}>
                            <TextField label="Subject" readOnly value={data.subject} />
                            {/* <TextField label="Call From" readOnly value={data.fromname} /> */}
                            <div style={{ marginBottom: 12 }}>
                                <Label>Call From</Label>
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
                                            openCrmForm(data.from_entitytype, data.from);
                                        }}
                                        style={{ textDecoration: "underline", color: "#0078d4" }}
                                    >
                                        {data.fromname}
                                    </a>
                                </div>
                            </div>
                            {/* <TextField label="Call To" readOnly value={data.toname} /> */}
                            <div style={{ marginBottom: 12 }}>
                                <Label>Call To</Label>
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
                                    {data.to && data.to.length > 0 ? (
                                        data.to.map((callto, index) => (
                                            <React.Fragment key={callto.partyid}>
                                                <a
                                                    href="#"
                                                    onClick={(e) => {
                                                        e.preventDefault();
                                                        openCrmForm(callto.entitytype, callto.partyid);
                                                    }}
                                                    style={{
                                                        textDecoration: "underline",
                                                        color: "#0078d4",
                                                        marginRight: "4px",
                                                    }}
                                                    title={callto.name}
                                                >
                                                    {callto.name}
                                                </a>
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
                            <TextField label="Phone Number" readOnly value={data.phonenumber} />
                        </Stack>
                        <Stack tokens={{ childrenGap: 10 }} styles={{ root: { border: "1px solid #e1dfdd", padding: "10px", borderRadius: "2px" } }}>
                            <TextField label="Phone Call Reason / Disposition" readOnly value={data.new_phonecallreason} />
                            <TextField label="Phone Call Outcome" readOnly value={data.new_phonecalloutcome} />
                            <TextField label="Note" readOnly value={data.strava_note} multiline/>
                            <TextField label="Note Category" readOnly value={data.strava_notecategory} />
                            <TextField label="Publish Note" readOnly value={data.strava_publishnote} />
                            <TextField label="Description" readOnly value={data.description} />
                        </Stack>
                        <Stack tokens={{ childrenGap: 10 }} styles={{ root: { border: "1px solid #e1dfdd", padding: "10px", borderRadius: "2px" } }}>
                            <TextField label="Queue Name" readOnly value={data.softphon_queuename} />
                            <TextField label="Queue time (seconds)" readOnly value={data.softphon_queuetimeseconds} />
                            <TextField label="Genesys Cloud WrapUpCode" readOnly value={data.softphon_genesyscloudwrapupcode} />
                            <TextField label="IVR time (seconds)" readOnly value={data.softphon_ivrtimeseconds} />
                            <TextField label="Duration" readOnly value={data.actualdurationminutes} />
                            <TextField label="Duration (seconds)" readOnly value={data.softphon_durationseconds} />
                            <TextField label="Disposition duration (seconds)" readOnly value={data.softphon_dispositiondurationseconds} />
                            <TextField label="Interaction URL" readOnly value={data.softphon_interactionurl} />
                        </Stack>
                    </Stack>
                    {/* Right Column */}
                    <Stack tokens={{ childrenGap: 15 }} styles={{ root: { flex: 6, minWidth: "400px"} }}>
                        <Stack tokens={{ childrenGap: 10 }} styles={{ root: { border: "1px solid #e1dfdd", padding: "10px", borderRadius: "2px" } }}>
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
                        <TextField label="Actual Start" readOnly value={data.actualstart} />
                        <TextField label="Actual End" readOnly value={data.actualend} />
                        <TextField label="Direction" readOnly value={data.directioncode} /></Stack>
                    </Stack>
                </Stack>
            </Stack>
        </Modal>
    );
};

export default PhoneCallForm;
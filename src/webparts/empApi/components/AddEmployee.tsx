import * as React from "react";
import { Stack, PrimaryButton, MessageBar, MessageBarType, TextField } from "office-ui-fabric-react";
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import axios from "axios";
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { spfi } from "@pnp/sp";  // Import PnP for SharePoint interaction
import { SPFx } from "@pnp/sp";   // Import SPFx context for PnP JS
import "@pnp/sp/webs";           // Import PnP JS for SharePoint webs
import "@pnp/sp/folders";        // Import PnP JS for SharePoint folders
import "@pnp/sp/files";          // Import PnP JS for SharePoint files

const apiUrl = "https://localhost:7042/api/Employees/add"; // API endpoint

interface IAddEmployeeProps {
  context: WebPartContext;
}

interface IAddEmployeeState {
  employeeName: string;
  salary: string;
  createdBy: string;
  createdTs: string;
  modifiedBy: string;
  modifiedTs: string;
  flag: string;
  active: string;
  errorMessage: string;
  modifiedByEmail: string;
  createdByEmail: string;
  department: string;
  designation: string;
  contactInfo: string;
  address: string;
  dateOfJoining: string;
  assignedUsers: string;
  Document_Link: File | null; // New field for file upload
}

export default class AddEmployee extends React.Component<IAddEmployeeProps, IAddEmployeeState> {
  [x: string]: any;
  private sp = spfi().using(SPFx(this.props.context)); // Initialize PnP JS with SPFx context

  constructor(props: IAddEmployeeProps) {
    super(props);
    this.state = {
      employeeName: "",
      salary: "",
      createdBy: "",
      createdTs: new Date().toISOString(),
      modifiedBy: "",
      modifiedTs: new Date().toISOString(),
      flag: "",
      active: "true",
      errorMessage: "",
      modifiedByEmail: "",
      createdByEmail: "",
      department: "",
      designation: "",
      contactInfo: "",
      address: "",
      dateOfJoining: new Date().toISOString(),
      assignedUsers: "",
      Document_Link: null, // Initialize with null
    };

    this.peoplePickerContext = {
      absoluteUrl: this.props.context.pageContext.web.absoluteUrl,
      msGraphClientFactory: this.props.context.msGraphClientFactory,
      spHttpClient: this.props.context.spHttpClient
    };
  }

  handleChange = (field: keyof IAddEmployeeState, value: string) => {
    this.setState({ [field]: value } as unknown as Pick<IAddEmployeeState, keyof IAddEmployeeState>);
  };

  handleFileChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files ? event.target.files[0] : null;
    this.setState({ Document_Link: file });
  };

  // Function to upload file to SharePoint document library
  uploadFile = async (): Promise <string | null> => {
    const {Document_Link}=this.state
    const targetLibraryUrl = "supported_document"; // Replace with your SharePoint document library URL
    // const {file}=this.state
    if (Document_Link) {
      try {
        // Upload the file to the SharePoint folder
        const response = await this.sp.web.getFolderByServerRelativePath(targetLibraryUrl).files.addUsingPath(Document_Link.name, Document_Link, { Overwrite: true });
        console.log("File uploaded successfully to SharePoint!");
        return response.ServerRelativeUrl; // Return the uploaded file's URL
      } catch (error) {
        console.error("Error uploading file to SharePoint:", error);
        return null;
      }
    }
    return null
  };

  handleAddEmployee = async () => {
    const { employeeName, salary, createdBy, modifiedBy, flag, active, department, designation, contactInfo, address, dateOfJoining, assignedUsers } = this.state;

    if (!employeeName || !salary || !createdBy || !modifiedBy || !flag || !active || !department || !designation || !contactInfo || !address || !dateOfJoining || !assignedUsers) {
      this.setState({ errorMessage: "Please fill in all fields and select users for 'Created By' and 'Modified By'." });
      return;
    }

    // Upload file to SharePoint (without storing the link in the database)
   
      const doclink=await this.uploadFile();
      if(!doclink){
        this.setState({errorMessage:"Failed to upload file to SharePoint"});
        return;
      }

    const employeeData = {
      employeeName,
      salary,
      createdBy: this.state.createdByEmail,
      createdTs: this.state.createdTs,
      modifiedBy: this.state.modifiedByEmail,
      modifiedTs: this.state.modifiedTs,
      flag,
      active,
      department,
      designation,
      contactInfo,
      address,
      dateOfJoining,
      assignedUsers,
      Document_Link:doclink
      // Do not include documentLink in the request if you don't want to store it in the database
    };

    try {
      const response = await axios.post(apiUrl, employeeData, {
        headers: { "Content-Type": "application/json" },
      });

      console.log("Employee added successfully:", response.data);
      alert("Employee added successfully!");

      // Reset the form after submission
      this.setState({
        employeeName: "",
        salary: "",
        createdBy: "",
        modifiedBy: "",
        flag: "",
        active: "true",
        createdTs: new Date().toISOString(),
        modifiedTs: new Date().toISOString(),
        department: "",
        designation: "",
        contactInfo: "",
        address: "",
        dateOfJoining: new Date().toISOString(),
        assignedUsers: "",
        Document_Link: null,
        errorMessage: "",
        
      });

    } catch (error) {
      console.error("Error adding employee:", error);
      this.setState({ errorMessage: error.response?.data?.message || "Failed to add employee" });
    }
  };

  handlePeoplePickerChange = (field: keyof IAddEmployeeState, items: any[]) => {
    const selectedUser = items.length > 0 ? items[0].id : "";  
    this.setState({ [field]: selectedUser } as Pick<IAddEmployeeState, keyof IAddEmployeeState>);
    this.setState({ modifiedByEmail: items[0].secondaryText, createdByEmail: items[0].secondaryText });
  };

  render() {
    const { employeeName, salary, errorMessage, Document_Link } = this.state;

    return (
      <Stack tokens={{ childrenGap: 10 }} style={{ maxWidth: 400, margin: "auto" }}>
        <h3>Add Employee</h3>

        {errorMessage && <MessageBar messageBarType={MessageBarType.error}>{errorMessage}</MessageBar>}

        <TextField label="Name" value={employeeName} onChange={(e, val) => this.handleChange("employeeName", val || "")} />
        <TextField label="Salary" type="number" value={salary} onChange={(e, val) => this.handleChange("salary", val || "")} />
        <TextField label="Department" value={this.state.department} onChange={(e, val) => this.handleChange("department", val || "")} />
        <TextField label="Designation" value={this.state.designation} onChange={(e, val) => this.handleChange("designation", val || "")} />
        <TextField label="Contact Info" value={this.state.contactInfo} onChange={(e, val) => this.handleChange("contactInfo", val || "")} />
        <TextField label="Address" value={this.state.address} onChange={(e, val) => this.handleChange("address", val || "")} />
        <TextField label="Date of Joining" type="date" value={this.state.dateOfJoining} onChange={(e, val) => this.handleChange("dateOfJoining", val || "")} />
        <TextField label="Assigned Users" value={this.state.assignedUsers} onChange={(e, val) => this.handleChange("assignedUsers", val || "")} />

        {/* People Picker for Created By */}
        <PeoplePicker
          context={this.peoplePickerContext}
          titleText="Created By"
          personSelectionLimit={1}
          groupName={""}
          showtooltip={true}
          required={true}
          ensureUser={true}
          defaultSelectedUsers={[this.state.createdBy]}
          onChange={(items) => this.handlePeoplePickerChange("createdBy", items)}
          principalTypes={[PrincipalType.User]}
          resolveDelay={1000}
        />

        {/* People Picker for Modified By */}
        <PeoplePicker
          context={this.peoplePickerContext}
          titleText="Modified By"
          personSelectionLimit={1}
          groupName={""}
          showtooltip={true}
          required={true}
          ensureUser={true}
          defaultSelectedUsers={[this.state.modifiedBy]}
          onChange={(items) => this.handlePeoplePickerChange("modifiedBy", items)}
          showHiddenInUI={false}
          principalTypes={[PrincipalType.User]}
          resolveDelay={1000}
        />

        <TextField label="Flag" value={this.state.flag} onChange={(e, val) => this.handleChange("flag", val || "")} />
        <TextField label="Active" value={this.state.active} onChange={(e, val) => this.handleChange("active", val || "")} />
        
        {/* Supporting Document Upload */}
        <label>Supporting Document:</label>
        <input type="file" accept=".pdf,.doc,.docx,.jpg,.png" onChange={this.handleFileChange} />
        {Document_Link && <p>Selected File: {Document_Link.name}</p>}

        <PrimaryButton onClick={this.handleAddEmployee} text="Add Employee" />
      </Stack>
    );
  }
}

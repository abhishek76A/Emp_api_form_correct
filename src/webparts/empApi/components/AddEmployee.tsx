import * as React from "react";
import { Stack, PrimaryButton, MessageBar, MessageBarType, TextField } from "office-ui-fabric-react";
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import axios from "axios";
import { WebPartContext } from '@microsoft/sp-webpart-base';

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
  supportingDocument: File | null; // New field for file upload
}

export default class AddEmployee extends React.Component<IAddEmployeeProps, IAddEmployeeState> {
  [x: string]: any;

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
      supportingDocument: null, // Initialize with null
    };

    // Corrected: Move peoplePickerContext inside the constructor
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
    this.setState({ supportingDocument: file });
  };

  handleAddEmployee = async () => {
    const { employeeName, salary, createdBy, modifiedBy, flag, active, department, designation, contactInfo, address, dateOfJoining, assignedUsers, supportingDocument } = this.state;

    if (!employeeName || !salary || !createdBy || !modifiedBy || !flag || !active || !department || !designation || !contactInfo || !address || !dateOfJoining || !assignedUsers) {
      this.setState({ errorMessage: "Please fill in all fields and select users for 'Created By' and 'Modified By'." });
      return;
    }

    const formData = new FormData();
    formData.append("name", employeeName);
    formData.append("salary", salary);
    formData.append("created_by", this.state.createdByEmail);
    formData.append("created_ts", this.state.createdTs);
    formData.append("modified_by", this.state.modifiedByEmail);
    formData.append("modified_ts", this.state.modifiedTs);
    formData.append("flag", flag);
    formData.append("active", active);
    formData.append("department", department);
    formData.append("designation", designation);
    formData.append("contactInfo", contactInfo);
    formData.append("address", address);
    formData.append("dateOfJoining", dateOfJoining);
    formData.append("assignedUsers", assignedUsers);

    if (supportingDocument) {
      formData.append("supportingDocument", supportingDocument); // Append the file
    }

    console.log("Sending Employee Data:", formData);

    try {
      const response = await axios.post(apiUrl, formData, {
        headers: { "Content-Type": "multipart/form-data" },
      });

      console.log("Employee added successfully:", response.data);
      alert("Employee added successfully!");

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
        supportingDocument: null,
        errorMessage: ""
      });

    } catch (error) {
      console.error("Error adding employee:", error);
      this.setState({ errorMessage: error.response?.data?.message || "Failed to add employee" });
    }
  };

  handlePeoplePickerChange = (field: keyof IAddEmployeeState, items: any[]) => {
    console.log(items);
    const selectedUser = items.length > 0 ? items[0].id : "";  
    this.setState({ [field]: selectedUser } as Pick<IAddEmployeeState, keyof IAddEmployeeState>);
    this.setState({ modifiedByEmail: items[0].secondaryText, createdByEmail: items[0].secondaryText });
  };

  render() {
    const { employeeName, salary, errorMessage, supportingDocument } = this.state;

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
        {supportingDocument && <p>Selected File: {supportingDocument.name}</p>}

        <PrimaryButton onClick={this.handleAddEmployee} text="Add Employee" />
      </Stack>
    );
  }
}

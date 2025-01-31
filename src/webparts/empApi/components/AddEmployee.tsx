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
      createdByEmail: ""
    };

    // Corrected: Move peoplePickerContext inside the constructor
    this.peoplePickerContext = {
      absoluteUrl: this.props.context.pageContext.web.absoluteUrl,
      msGraphClientFactory: this.props.context.msGraphClientFactory,
      spHttpClient: this.props.context.spHttpClient
    };
  }

  handleChange = (field: keyof IAddEmployeeState, value: string) => {
    this.setState({ [field]: value } as Pick<IAddEmployeeState, keyof IAddEmployeeState>);
  };

  handleAddEmployee = async () => {
    const { employeeName, salary, createdBy, modifiedBy, flag, active } = this.state;

    // Ensure createdBy and modifiedBy are strings before trimming
    if (!employeeName || !salary || !createdBy || !modifiedBy || !flag || !active) {
      this.setState({ errorMessage: "Please fill in all fields and select users for 'Created By' and 'Modified By'." });
      return;
    }

    const employeeData = {
      id: 0,
      name: employeeName,
      salary: parseFloat(salary),
      created_by: this.state.createdByEmail,  // Ensure this is a string
      created_ts: this.state.createdTs,
      modified_by: this.state.modifiedByEmail,  // Ensure this is a string
      modified_ts: this.state.modifiedTs,
      flag: flag,
      active: active,
    };

    console.log("Sending Employee Data:", employeeData);
    try {
      const response = await axios.post(apiUrl, employeeData, {
        headers: { "Content-Type": "application/json" },
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
        errorMessage: ""
      });

    } catch (error) {
      console.error("Error adding employee:", error);
      this.setState({ errorMessage: error.response?.data?.message || "Failed to add employee" });
    }
  };

  handlePeoplePickerChange = (field: keyof IAddEmployeeState, items: any[]) => {
    console.log(items);
    // Ensure that we're extracting a string (user ID or display name)
    const selectedUser = items.length > 0 ? items[0].id : "";  // Extract the user ID (string) from the PeoplePicker result
    this.setState({ [field]: selectedUser } as Pick<IAddEmployeeState, keyof IAddEmployeeState>);
    this.setState({modifiedByEmail : items[0].secondaryText});
    this.setState({createdByEmail : items[0].secondaryText});
  };

  render() {
    const { employeeName, salary, errorMessage } = this.state;

    return (
      <Stack tokens={{ childrenGap: 10 }} style={{ maxWidth: 400, margin: "auto" }}>
        <h3>Add Employee</h3>

        {errorMessage && <MessageBar messageBarType={MessageBarType.error}>{errorMessage}</MessageBar>}

        <TextField label="Name" value={employeeName}
          onChange={(e, val) => this.handleChange("employeeName", val || "")} />

        <TextField label="Salary" type="number" value={salary}
          onChange={(e, val) => this.handleChange("salary", val || "")} />

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

        <TextField label="Flag" value={this.state.flag}
          onChange={(e, val) => this.handleChange("flag", val || "")} />

        <TextField label="Active" value={this.state.active}
          onChange={(e, val) => this.handleChange("active", val || "")} />

        <PrimaryButton onClick={this.handleAddEmployee} text="Add Employee" />
      </Stack>
    );
  }
}

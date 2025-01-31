import * as React from "react";
import { PrimaryButton, TextField, MessageBar, MessageBarType, Stack, Spinner, SpinnerSize } from "office-ui-fabric-react";
import axios from "axios";

const apiUrl = "https://localhost:7042/api/Employees";

interface UpdateEmployeeByIdProps {
  updateEmployeeData: (data: any) => void;
}

interface EmployeeState {
  employeeId: string;
  employee: any | null;
  loading: boolean;
  error: string | null;
  updatedEmployee: any | null;
}

class UpdateEmployeeById extends React.Component<UpdateEmployeeByIdProps, EmployeeState> {
  constructor(props: UpdateEmployeeByIdProps) {
    super(props);
    this.state = {
      employeeId: "",
      employee: null,
      loading: false,
      error: null,
      updatedEmployee: null
    };
  }

  componentDidMount() {
    // You can add any initialization logic here if needed
  }

  fetchEmployeeDetails = async () => {
    if (!this.state.employeeId) {
      this.setState({ error: "Please enter an Employee ID." });
      return;
    }

    this.setState({ loading: true, error: null });

    try {
      const response = await axios.get(`${apiUrl}/details/${this.state.employeeId}`);
      if (response.data) {
        this.setState({
          employee: response.data,
          updatedEmployee: response.data,
          error: null
        });
      } else {
        this.setState({ error: "Employee not found.", employee: null });
      }
    } catch (error) {
      console.error("Error fetching employee details:", error);
      this.setState({ 
        error: "An error occurred while fetching employee details. Please try again.",
        employee: null
      });
    } finally {
      this.setState({ loading: false });
    }
  };

  updateEmployeeDetails = async () => {
    if (!this.state.updatedEmployee) {
      this.setState({ error: "No employee data to update." });
      return;
    }

    const updatedEmployeeWithModifiedBy = {
      ...this.state.updatedEmployee,
      Modified_by: "System",
      Modified_ts: new Date().toISOString()
    };

    this.setState({ loading: true, error: null });

    try {
      const response = await axios.put(`${apiUrl}/update/${this.state.employeeId}`, updatedEmployeeWithModifiedBy);
      if (response.data) {
        this.props.updateEmployeeData(updatedEmployeeWithModifiedBy);
        this.setState({
          employee: updatedEmployeeWithModifiedBy,
          error: null
        });
        alert("Employee updated successfully!");
      }
    } catch (error) {
      console.error("Error updating employee details:", error);
      this.setState({ error: "Failed to update employee details. Please try again." });
    } finally {
      this.setState({ loading: false });
    }
  };

  handleChange = (field: string, value: any) => {
    this.setState(prevState => ({
      updatedEmployee: {
        ...prevState.updatedEmployee,
        [field]: value
      }
    }));
  };

  render() {
    const { employeeId, employee, loading, error, updatedEmployee } = this.state;

    return (
      <Stack tokens={{ childrenGap: 20 }} style={{ padding: 20 }}>
        <h3>Update Employee Details by ID</h3>

        <TextField
          label="Employee ID"
          value={employeeId}
          onChange={(e, newValue) => this.setState({ employeeId: newValue || "" })}
        />
        <PrimaryButton
          onClick={this.fetchEmployeeDetails}
          text="Fetch Employee"
          disabled={loading}
        />

        {loading && <Spinner size={SpinnerSize.medium} label="Fetching..." />}
        {error && <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>}

        {employee && !loading && !error && (
          <div>
            <TextField
              label="Name"
              value={updatedEmployee?.name || ""}
              onChange={(e, newValue) => this.handleChange("name", newValue)}
            />
            <TextField
              label="Salary"
              type="number"
              value={updatedEmployee?.salary || ""}
              onChange={(e, newValue) => this.handleChange("salary", newValue)}
            />
            <TextField
              label="Created By"
              value={updatedEmployee?.created_by || ""}
              onChange={(e, newValue) => this.handleChange("created_by", newValue)}
            />
            <TextField
              label="Flag"
              value={updatedEmployee?.flag || ""}
              onChange={(e, newValue) => this.handleChange("flag", newValue)}
            />
            <TextField
              label="Active"
              value={updatedEmployee?.active || ""}
              onChange={(e, newValue) => this.handleChange("active", newValue)}
            />
            <PrimaryButton
              onClick={this.updateEmployeeDetails}
              text="Update Employee Details"
              disabled={loading}
            />
          </div>
        )}

        {employee === null && !loading && !error && (
          <MessageBar>No employee found for this ID.</MessageBar>
        )}
      </Stack>
    );
  }
}

export default UpdateEmployeeById;
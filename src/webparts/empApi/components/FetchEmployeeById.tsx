import * as React from "react";
import { PrimaryButton, TextField, MessageBar, MessageBarType, Stack, Spinner, SpinnerSize, DetailsList, DetailsListLayoutMode, IColumn } from "office-ui-fabric-react";
import axios from "axios";

const apiUrl = "https://localhost:7042/api/Employees";

interface IProps {
  updateEmployeeData: (data: any) => void;
}

interface IState {
  employeeId: string;
  employee: any | null;
  loading: boolean;
  error: string | null;
}

class FetchEmployeeById extends React.Component<IProps, IState> {
  constructor(props: IProps) {
    super(props);
    this.state = {
      employeeId: "",
      employee: null,
      loading: false,
      error: null
    };
  }

  componentDidMount() {
    // Initialization logic can be added here if needed
  }

  // Define columns as class property since they're static
  private columns: IColumn[] = [
    { key: "id", name: "ID", fieldName: "id", minWidth: 50, maxWidth: 100, isResizable: true },
    { key: "name", name: "Name", fieldName: "name", minWidth: 100, maxWidth: 200, isResizable: true },
    { key: "salary", name: "Salary", fieldName: "salary", minWidth: 100, maxWidth: 150, isResizable: true },
    { key: "created_by", name: "Created By", fieldName: "created_by", minWidth: 100, maxWidth: 200, isResizable: true },
    { key: "created_ts", name: "Created Timestamp", fieldName: "created_ts", minWidth: 150, maxWidth: 200, isResizable: true },
    { key: "modified_by", name: "Modified By", fieldName: "modified_by", minWidth: 100, maxWidth: 200, isResizable: true },
    { key: "modified_ts", name: "Modified Timestamp", fieldName: "modified_ts", minWidth: 150, maxWidth: 200, isResizable: true },
    { key: "flag", name: "Flag", fieldName: "flag", minWidth: 50, maxWidth: 100, isResizable: true },
    { key: "active", name: "Active", fieldName: "active", minWidth: 80, maxWidth: 100, isResizable: true },
  ];

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
          error: null
        });
        this.props.updateEmployeeData(response.data);
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

  render() {
    const { employeeId, employee, loading, error } = this.state;

    return (
      <Stack tokens={{ childrenGap: 20 }} style={{ padding: 20 }}>
        <h3>Fetch Employee Details by ID</h3>
        <TextField
          label="Employee ID"
          value={employeeId}
          onChange={(e, newValue) => this.setState({ employeeId: newValue || "" })}
        />
        <PrimaryButton
          onClick={this.fetchEmployeeDetails}
          text="Fetch Employee Details"
          disabled={loading}
        />

        {loading && <Spinner size={SpinnerSize.medium} label="Fetching..." />}
        {error && <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>}

        {employee && !loading && !error && (
          <DetailsList
            items={[employee]}
            columns={this.columns}
            setKey="set"
            layoutMode={DetailsListLayoutMode.justified}
          />
        )}

        {employee === null && !loading && !error && (
          <MessageBar>No employee found for this ID.</MessageBar>
        )}
      </Stack>
    );
  }
}

export default FetchEmployeeById;
import * as React from "react";
import { Stack, MessageBar, MessageBarType, Spinner, DetailsList, DetailsListLayoutMode, IColumn } from "office-ui-fabric-react";
import axios from "axios";

const apiUrl = "https://localhost:7042/api/Employees";

interface IState {
  employees: any[];
  loading: boolean;
  error: string | null;
}

class EmployeeList extends React.Component<{}, IState> {
  constructor(props: {}) {
    super(props);
    this.state = {
      employees: [],
      loading: true,
      error: null
    };
  }

  componentDidMount() {
    this.fetchEmployees();
  }

  // Define columns as class property
  private columns: IColumn[] = [
    { key: "id", name: "ID", fieldName: "id", minWidth: 50, maxWidth: 100, isResizable: true },
    { key: "name", name: "Name", fieldName: "name", minWidth: 100, maxWidth: 200, isResizable: true },
    { key: "salary", name: "Salary", fieldName: "salary", minWidth: 100, maxWidth: 150, isResizable: true },
    { key: "created_by", name: "Created By", fieldName: "created_by", minWidth: 100, maxWidth: 200, isResizable: true },
    { key: "active", name: "Active", fieldName: "active", minWidth: 80, maxWidth: 100, isResizable: true },
  ];

  fetchEmployees = async () => {
    try {
      const response = await axios.get(`${apiUrl}/all`);
      this.setState({
        employees: response.data,
        loading: false,
        error: null
      });
    } catch (error) {
      console.error("Error fetching employees:", error);
      this.setState({
        error: "Failed to load employee data. Please try again.",
        loading: false
      });
    }
  };

  render() {
    const { employees, loading, error } = this.state;

    return (
      <Stack tokens={{ childrenGap: 20 }} style={{ padding: 20 }}>
        <h3>Employee List</h3>

        {/* State Management Indicators */}
        {loading && <Spinner label="Loading employees..." />}
        {error && <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>}
        {employees.length === 0 && !loading && !error && (
          <MessageBar>No employees found.</MessageBar>
        )}

        {/* Employee Data Table */}
        {employees.length > 0 && !loading && !error && (
          <DetailsList
            items={employees}
            columns={this.columns}
            setKey="set"
            layoutMode={DetailsListLayoutMode.justified}
            selectionPreservedOnEmptyClick
          />
        )}
      </Stack>
    );
  }
}

export default EmployeeList;
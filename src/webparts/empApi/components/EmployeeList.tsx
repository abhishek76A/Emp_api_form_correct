import * as React from "react";
import { Stack, MessageBar, MessageBarType, Spinner, DetailsList, DetailsListLayoutMode, IColumn, DefaultButton } from "office-ui-fabric-react";
import axios from "axios";

const apiUrl = "https://localhost:7042/api/Employees";

interface IState {
  employees: any[];
  loading: boolean;
  error: string | null;
  message: string | null;
}

class EmployeeList extends React.Component<{}, IState> {
  constructor(props: {}) {
    super(props);
    this.state = {
      employees: [],
      loading: true,
      error: null,
      message: null
    };
  }

  componentDidMount() {
    this.fetchEmployees();
  }

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

  handleDelete = async (id: number) => {
    console.log("Deleting employee with ID:", id);
    if (!window.confirm(`Are you sure you want to delete Employee ID: ${id}?`)) return;

    try {
      await axios.delete(`${apiUrl}/remove/${id}`);
      this.setState(prevState => ({
        employees: prevState.employees.filter(emp => emp.id !== id),
        message: `✅ Employee with ID ${id} has been deleted successfully.`
      }));
    } catch (error) {
      console.error("Error deleting employee:", error);
      this.setState({ message: "❌ Failed to delete employee. Please try again." });
    }
  };

  private columns: IColumn[] = [
    { key: "id", name: "ID", fieldName: "id", minWidth: 50, maxWidth: 100, isResizable: true },
    { key: "name", name: "Name", fieldName: "name", minWidth: 100, maxWidth: 200, isResizable: true },
    { key: "salary", name: "Salary", fieldName: "salary", minWidth: 100, maxWidth: 150, isResizable: true },
    { key: "department", name: "Department", fieldName: "department", minWidth: 100, maxWidth: 200, isResizable: true },
    { key: "designation", name: "Designation", fieldName: "designation", minWidth: 100, maxWidth: 200, isResizable: true },
    { key: "address", name: "Address", fieldName: "address", minWidth: 150, maxWidth: 300, isResizable: true },
    {
      key: "delete",
      name: "Actions",
      minWidth: 100,
      maxWidth: 150,
      isResizable: false,
      onRender: (item) => (
        <DefaultButton 
          text="Delete" 
          onClick={() => this.handleDelete(item.id)} 
          style={{ backgroundColor: "red", color: "white" }} 
        />
      ),
    },
  ];

  render() {
    const { employees, loading, error, message } = this.state;

    return (
      <Stack tokens={{ childrenGap: 20 }} style={{ padding: 20 }}>
        <h3>Employee List</h3>

        {/* State Management Indicators */}
        {loading && <Spinner label="Loading employees..." />}
        {error && <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>}
        {message && <MessageBar messageBarType={MessageBarType.warning}>{message}</MessageBar>}
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

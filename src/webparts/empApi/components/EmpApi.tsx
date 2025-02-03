import * as React from "react";
import AddEmployee from "./AddEmployee";
import FetchEmployeeById from "./FetchEmployeeById";
import EmployeeList from "./EmployeeList";
import { IEmpApiProps } from "./IEmpApiProps";

// Define your inline button styles with smaller buttons
const buttonStyle = {
  backgroundColor: '#0078D4', // Blue background
  color: 'white', // White text
  padding: '5px 10px', // Smaller padding inside the button
  fontSize: '14px', // Smaller font size
  border: 'none', // No border
  borderRadius: '5px', // Rounded corners
  cursor: 'pointer', // Pointer cursor on hover
  margin: '5px', // Smaller spacing between buttons
  transition: 'background-color 0.3s ease', // Smooth background color transition
};

const EmpApi: React.FC<IEmpApiProps> = (props) => {
  const { context } = props; // Ensure correct prop destructuring

  // Define a state to track the current view
  const [currentView, setCurrentView] = React.useState<string>("home");

  // Function to handle the button clicks and set the view
  const handleViewChange = (view: string): void => {
    setCurrentView(view); // Update the state to show the correct component
  };

  // Function to handle fetched employee data
  const updateEmployeeData = (data: any) => {
    console.log("Fetched Employee Data:", data);
  };

  return (
    <div>
      {/* Buttons to change the view with smaller inline CSS */}
      <button style={buttonStyle} onClick={() => handleViewChange("add")}>Add Employee</button>
      <button style={buttonStyle} onClick={() => handleViewChange("view")}>Employee List</button>
      <button style={buttonStyle} onClick={() => handleViewChange("id")}>Fetch Employee by ID</button>
      <button style={buttonStyle} onClick={() => handleViewChange("delete")}>Delete Employee</button>

      {/* Conditional rendering based on the current view */}
      <div>
        {currentView === "home" && <h3>Welcome to Employee Management System</h3>}
        {currentView === "add" && <AddEmployee context={context} />}
        {currentView === "view" && <EmployeeList />}
        {currentView === "id" && <FetchEmployeeById updateEmployeeData={updateEmployeeData} />}
      </div>
    </div>
  );
};

export default EmpApi;

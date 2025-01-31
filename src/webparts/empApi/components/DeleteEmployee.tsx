import * as React from "react";
import { useState } from "react";
import { TextField, PrimaryButton, Stack, MessageBar, MessageBarType } from "office-ui-fabric-react";
import axios from "axios";

const apiUrl = "https://localhost:7042/api/Employees";

const DeleteEmployee: React.FC = () => {
  const [employeeId, setEmployeeId] = useState<string>("");
  const [message, setMessage] = useState<string | null>(null);
  const [isLoading, setIsLoading] = useState<boolean>(false); // Prevent multiple clicks

  // Function to delete employee
  const handleDelete = async () => {
    setMessage(null); // Clear previous messages
    if (!employeeId) return setMessage("❌ Please enter a valid Employee ID.");

    if (!window.confirm(`Are you sure you want to delete Employee ID: ${employeeId}?`)) return;

    try {
      setIsLoading(true); // Disable button while deleting
      await axios.delete(`${apiUrl}/remove/${employeeId}`);
      setMessage(`✅ Employee with ID ${employeeId} has been deleted successfully.`);
      setEmployeeId(""); // Reset input field
    } catch (error) {
      console.error("Error deleting employee:", error);
      setMessage("❌ Failed to delete employee. Please try again.");
    } finally {
      setIsLoading(false);
    }
  };

  return (
    <Stack tokens={{ childrenGap: 10 }} style={{ padding: 20 }}>
      <h3>Delete Employee</h3>

      {/* Success/Error Message */}
      {message && <MessageBar messageBarType={MessageBarType.warning}>{message}</MessageBar>}

      <TextField 
        label="Employee ID" 
        value={employeeId} 
        onChange={(e, val) => setEmployeeId(val || "")} 
      />

      <PrimaryButton 
        onClick={handleDelete} 
        text={isLoading ? "Deleting..." : "Delete Employee"} 
        disabled={isLoading} 
        style={{ backgroundColor: "red", color: "white" }} 
      />
    </Stack>
  );
};

export default DeleteEmployee;

import * as React from "react";
import { useState, useEffect } from "react";
import { IEmployee } from "./IEmployee";
import { IEmployeeWebPartProps } from "./IEmployeeWebPartProps";
import {
  Spinner,
  SpinnerSize,
  MessageBar,
  MessageBarType,
  DetailsList,
  DetailsListLayoutMode,
  IColumn,
  Selection,
  IObjectWithKey,
  TextField,
} from "@fluentui/react";
import EmployeeForm from "./EmployeeForm";
import { fetchEmployees, deleteEmployee } from "../actions";
import styles from "./EmployeeList.module.scss";

const columns: IColumn[] = [
  {
    key: "colNo",
    name: "No",
    fieldName: "No",
    minWidth: 30,
    maxWidth: 50,
    isResizable: true,
    onRender: (_, index?: number) => (index !== undefined ? index + 1 : 0),
  },
  {
    key: "colName",
    name: "Name",
    fieldName: "Name",
    minWidth: 100,
    maxWidth: 200,
    isResizable: true,
  },
  {
    key: "colJob",
    name: "Job Description",
    fieldName: "JobDescription",
    minWidth: 150,
    maxWidth: 300,
    isResizable: true,
  },
  {
    key: "colHired",
    name: "Hired On",
    fieldName: "HireDate",
    minWidth: 100,
    maxWidth: 150,
    isResizable: true,
    onRender: (item: IEmployee) => new Date(item.HireDate).toLocaleDateString(),
  },
];

const EmployeeList: React.FC<IEmployeeWebPartProps> = ({ context }) => {
  const [employees, setEmployees] = useState<IEmployee[]>([]);
  const [filteredEmployees, setFilteredEmployees] = useState<IEmployee[]>([]);
  const [fetching, setFetching] = useState<boolean>(true);
  const [fetchError, setFetchError] = useState<string | null>(null);
  const [selectedEmp, setSelectedEmp] = useState<IEmployee | undefined>(
    undefined
  );
  const [searchTerm, setSearchTerm] = useState("");

  const selection = new Selection({
    onSelectionChanged: () => {
      const sel: IObjectWithKey[] = selection.getSelection();
      setSelectedEmp(sel.length > 0 ? (sel[0] as IEmployee) : undefined);
    },
  });

  useEffect(() => {
    const loadEmployees = async (): Promise<void> => {
      try {
        setFetching(true);
        const data = await fetchEmployees(context);
        setEmployees(data);
        setFilteredEmployees(data);
      } catch (err) {
        setFetchError(err instanceof Error ? err.message : String(err));
      } finally {
        setFetching(false);
      }
    };
    loadEmployees().catch((err) => {
      setFetchError(err instanceof Error ? err.message : String(err));
      setFetching(false);
    });
  }, [context]);

  useEffect(() => {
    const filtered = employees.filter((emp) =>
      emp.Name.toLowerCase().includes(searchTerm.toLowerCase())
    );
    setFilteredEmployees(filtered);
  }, [searchTerm, employees]);

  const handleSuccess = async (): Promise<void> => {
    const data = await fetchEmployees(context);
    setEmployees(data);
    setFilteredEmployees(data);
    setSelectedEmp(undefined);
  };

  const handleDelete = async (): Promise<void> => {
    if (selectedEmp && selectedEmp.Id !== null) {
      await deleteEmployee(context, selectedEmp.Id);
      await handleSuccess();
    }
  };

  if (fetching) {
    return <Spinner label="Loading employees…" size={SpinnerSize.large} />;
  }

  if (fetchError) {
    return (
      <MessageBar messageBarType={MessageBarType.error}>
        Error loading employees: {fetchError}
      </MessageBar>
    );
  }

  if (employees.length === 0) {
    return (
      <MessageBar messageBarType={MessageBarType.warning}>
        No employees found.
      </MessageBar>
    );
  }

  return (
    <div>
      <h2>Employee Details</h2>

      <div className={styles.searchFieldContainer}>
        <p>select a row to edit or delete</p>

        <TextField
          placeholder="Search by name..."
          value={searchTerm}
          onChange={(e, value) => setSearchTerm(value || "")}
          className={styles.searchField}
        />
      </div>

      <DetailsList
        items={filteredEmployees}
        columns={columns}
        layoutMode={DetailsListLayoutMode.fixedColumns}
        isHeaderVisible
        selection={selection}
        selectionMode={1}
        className={styles.employeeList}
      />
      <EmployeeForm
        context={context}
        employee={selectedEmp}
        onSuccess={handleSuccess}
        onDelete={handleDelete}
      />
    </div>
  );
};

export default EmployeeList;

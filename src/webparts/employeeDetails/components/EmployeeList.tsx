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
    onRender: (_, index?: number) => index ?? 0,
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
  const [fetching, setFetching] = useState<boolean>(true);
  const [fetchError, setFetchError] = useState<string | null>(null);
  const [selectedEmp, setSelectedEmp] = useState<IEmployee | undefined>(
    undefined
  );

  const selection = new Selection({
    onSelectionChanged: () => {
      const sel: IObjectWithKey[] = selection.getSelection();
      if (sel.length > 0) {
        setSelectedEmp(sel[0] as IEmployee);
      } else {
        setSelectedEmp(undefined);
      }
    },
  });

  const loadEmployees = async (): Promise<void> => {
    setFetching(true);
    setFetchError(null);
    try {
      const data = await fetchEmployees(context);
      setEmployees(data);
    } catch (err: unknown) {
      setFetchError(err instanceof Error ? err.message : String(err));
    } finally {
      setFetching(false);
    }
  };

  useEffect(() => {
    loadEmployees().catch((err) => {
      console.error("Error loading employees:", err);
      setFetchError(err instanceof Error ? err.message : String(err));
    });
  }, [context]);

  const handleSuccess = async (): Promise<void> => {
    await loadEmployees();
    setSelectedEmp(undefined);
  };

  const handleDelete = async (): Promise<void> => {
    if (!selectedEmp || selectedEmp.Id === null) return;
    await deleteEmployee(context, selectedEmp.Id);
    await loadEmployees();
    setSelectedEmp(undefined);
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
      <DetailsList
        items={employees}
        columns={columns}
        layoutMode={DetailsListLayoutMode.fixedColumns}
        isHeaderVisible
        selection={selection}
        selectionMode={1} // single
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

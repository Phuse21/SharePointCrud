import * as React from "react";
import {
  TextField,
  DatePicker,
  Stack,
  MessageBar,
  MessageBarType,
  Dialog,
  DialogType,
  DialogFooter,
  PrimaryButton,
  DefaultButton,
} from "@fluentui/react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IEmployee } from "./IEmployee";
import Buttons from "./Buttons";
import { addEmployee, updateEmployee } from "../actions";

interface EmployeeFormProps {
  context: WebPartContext;
  employee?: IEmployee;
  onSuccess: () => void;
  onDelete?: () => Promise<void>;
  onCancel?: () => void;
}

const EmployeeForm: React.FC<EmployeeFormProps> = ({
  context,
  employee,
  onSuccess,
  onDelete,
  onCancel,
}) => {
  const [formData, setFormData] = React.useState({
    id: employee?.Id,
    name: employee?.Name ?? "",
    hireDate: employee?.HireDate ?? "",
    jobDescription: employee?.JobDescription ?? "",
  });
  const [errors, setErrors] = React.useState({
    name: "",
    hireDate: "",
    jobDescription: "",
  });
  const [saving, setSaving] = React.useState(false);
  const [error, setError] = React.useState<string | null>(null);
  const [confirmAction, setConfirmAction] = React.useState<
    "save" | "delete" | null
  >(null);

  React.useEffect(() => {
    if (employee) {
      setFormData({
        id: employee.Id,
        name: employee.Name,
        hireDate: employee.HireDate,
        jobDescription: employee.JobDescription,
      });
    } else {
      setFormData({
        id: undefined,
        name: "",
        hireDate: "",
        jobDescription: "",
      });
    }
    setErrors({ name: "", hireDate: "", jobDescription: "" });
    setError(null);
    setConfirmAction(null);
  }, [employee]);

  const validate = (): boolean => {
    const newErrors = {
      name: formData.name ? "" : "Name is required.",
      hireDate: formData.hireDate ? "" : "Hire Date is required.",
      jobDescription: formData.jobDescription
        ? ""
        : "Job Description is required.",
    };
    setErrors(newErrors);
    return !Object.values(newErrors).some((msg) => msg);
  };

  const handleSave = async (): Promise<void> => {
    setError(null);
    if (!validate()) return;
    setSaving(true);
    try {
      if (formData.id) {
        await updateEmployee(context, formData.id, {
          Name: formData.name,
          HireDate: new Date(formData.hireDate).toISOString(),
          JobDescription: formData.jobDescription,
        });
      } else {
        await addEmployee(context, {
          Name: formData.name,
          HireDate: new Date(formData.hireDate).toISOString(),
          JobDescription: formData.jobDescription,
        });
      }
      onSuccess();
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
    } finally {
      setSaving(false);
      setConfirmAction(null);
    }
  };

  const handleDeleteConfirm = async (): Promise<void> => {
    if (onDelete) {
      setSaving(true);
      await onDelete();
      setSaving(false);
      setConfirmAction(null);
    }
  };

  return (
    <>
      <form
        onSubmit={(e) => {
          e.preventDefault();
          setConfirmAction("save");
        }}
      >
        <Stack
          tokens={{ childrenGap: 15 }}
          styles={{ root: { width: "100%" } }}
        >
          {error && (
            <MessageBar messageBarType={MessageBarType.error}>
              {error}
            </MessageBar>
          )}
          <TextField
            label="Name"
            value={formData.name}
            onChange={(_, v) => {
              setFormData((fd) => ({ ...fd, name: v || "" }));
              setErrors((e) => ({ ...e, name: "" }));
            }}
            required
            errorMessage={errors.name}
          />
          <DatePicker
            label="Hire Date *"
            value={formData.hireDate ? new Date(formData.hireDate) : undefined}
            onSelectDate={(d) => {
              setFormData((fd) => ({
                ...fd,
                hireDate: d?.toISOString() || "",
              }));
              setErrors((e) => ({ ...e, hireDate: "" }));
            }}
            textField={{ errorMessage: errors.hireDate }}
          />
          <TextField
            label="Job Description"
            multiline
            rows={4}
            value={formData.jobDescription}
            onChange={(_, v) => {
              setFormData((fd) => ({ ...fd, jobDescription: v || "" }));
              setErrors((e) => ({ ...e, jobDescription: "" }));
            }}
            required
            errorMessage={errors.jobDescription}
          />
        </Stack>
        <Buttons
          onSave={async (): Promise<void> => setConfirmAction("save")}
          saving={saving}
          isEditing={!!employee?.Id}
          onDelete={async (): Promise<void> => setConfirmAction("delete")}
          onCancel={onCancel}
        />
      </form>

      <Dialog
        hidden={!confirmAction}
        onDismiss={() => setConfirmAction(null)}
        dialogContentProps={{
          type: DialogType.normal,
          title: confirmAction === "delete" ? "Confirm Delete" : "Confirm Save",
          subText:
            confirmAction === "delete"
              ? "Are you sure you want to delete this employee?"
              : "Do you want to save the changes?",
        }}
      >
        <DialogFooter>
          <PrimaryButton
            onClick={
              confirmAction === "delete" ? handleDeleteConfirm : handleSave
            }
            text={confirmAction === "delete" ? "Yes, delete" : "Yes, save"}
          />
          <DefaultButton onClick={() => setConfirmAction(null)} text="No" />
        </DialogFooter>
      </Dialog>
    </>
  );
};

export default EmployeeForm;

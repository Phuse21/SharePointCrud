import * as React from "react";
import {
  DefaultButton,
  PrimaryButton,
  Spinner,
  SpinnerSize,
} from "@fluentui/react";
import styles from "./EmployeeList.module.scss";

interface ButtonsProps {
  /** Called for both create and update actions */
  onSave: () => Promise<void>;
  /** True while the save (create/update) is in progress */
  saving: boolean;
  /** Indicates whether we're editing an existing record */
  isEditing?: boolean;
  /** Optional delete callback; shown only when editing */
  onDelete?: () => void;
  /** Optional cancel callback; shown only when editing or deleting */
  onCancel?: () => void;
}

const Buttons: React.FC<ButtonsProps> = ({
  onSave,
  saving,
  isEditing = false,
  onDelete,
  onCancel,
}) => {
  const primaryText = saving
    ? isEditing
      ? "Updating…"
      : "Creating…"
    : isEditing
    ? "Update"
    : "Create";

  return (
    <div className={styles.buttonContainer}>
      <PrimaryButton
        text={primaryText}
        onClick={onSave}
        disabled={saving}
        type="button"
      >
        {saving && (
          <Spinner
            size={SpinnerSize.small}
            label=""
            ariaLive="assertive"
            styles={{ root: { marginLeft: 8 } }}
          />
        )}
      </PrimaryButton>

      {isEditing && onDelete && (
        <DefaultButton
          text="Delete"
          onClick={onDelete}
          className={styles.deleteButton}
          type="button"
        />
      )}

      {isEditing && onCancel && (
        <DefaultButton text="Cancel" onClick={onCancel} type="button" />
      )}
    </div>
  );
};

export default Buttons;

# Disabling Source & Destination Changes During Cable Edit

When users edit an existing cable, the register form (`frm_RegisterCable`) should prevent
any modifications to the **Source** and **Destination** endpoint selectors. If the
comboboxes are left enabled while their lookup fails (for example, due to missing endpoint
records), the controls remain blank and saving the record overwrites the stored endpoints
with empty values.

To avoid this data loss scenario, ensure that the form automatically locks these controls
whenever it is launched in **UPDATE** mode. The change is implemented entirely inside the
form's code-behind module.

## Implementation

1. In the VBA editor, open the project tree `Microsoft Excel Objects → Forms → frm_RegisterCable`.
2. Locate the routine that prepares the form for editing an existing cable. In the stock
workbook this happens inside `frm_RegisterCable.LoadCableForEdit`.
3. After the routine has populated all of the form fields, add:

   ```vb
   Me.cmb_Source.Enabled = False
   Me.cmb_Destination.Enabled = False
   Me.cmb_Source.Locked = True
   Me.cmb_Destination.Locked = True
   ```

   Locking as well as disabling guarantees that keyboard focus or code cannot alter the
   values while the user reviews the record.
4. If the form later switches back to **CREATE** mode (for example, when the user chooses
   “Save & Continue”), re-enable both controls so that new cables can set their endpoints:

   ```vb
   Me.cmb_Source.Enabled = True
   Me.cmb_Destination.Enabled = True
   Me.cmb_Source.Locked = False
   Me.cmb_Destination.Locked = False
   ```
5. Finally, confirm that any validation routine that raises a “missing source/destination”
   error skips these fields when the form is in update mode, since the saved data remains in
   the worksheet rather than in the disabled controls.

## Testing Checklist

- Launch an edit action from the register table and verify that both comboboxes are greyed
  out while still displaying the stored endpoint names.
- Attempt to save without touching the controls. The record should retain its original
  source and destination values.
- Start a new-cable registration; the comboboxes must once again be active and allow the
  selection of endpoints.

Following these steps keeps the workbook from clearing endpoint data during cable edits
while preserving the ability to configure new cables normally.
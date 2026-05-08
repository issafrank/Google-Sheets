# Inventory & Maintenance — Radio Workbook (Google Apps Script)

Apps Script para sa Google Sheets na nag-bu-build ng inventory + maintenance + issuance + audit + return forms para sa radio assets.

## Setup
1. Buksan ang target Google Sheet.
2. Pumunta sa **Extensions → Apps Script**.
3. I-paste ang buong laman ng `Code.gs` sa script editor.
4. I-save (Ctrl+S), tapos i-reload ang spreadsheet.
5. Sa menu bar, i-click ang **Inventory & Maintenance → Setup Radio Workbook**.

## Menu
- **Setup Radio Workbook** — gumagawa/nagre-rebuild ng lahat ng sheets at forms (preserves existing data).
- **Apply Condition Colors** — nag-aapply lang ng conditional formatting sa `ASSET_MASTER`.
- **Refresh Validations Only** — nire-refresh ang dropdowns at flow validations.

## Sheets na ginagawa
- `AssetMasterForm`, `RadioIssuanceForm`, `MaintenanceEntryForm`, `InventoryAuditForm`, `RadioReturnForm` — input forms.
- `ASSET_MASTER`, `ISSUANCE_LOG`, `MAINTENANCE_LOG`, `INVENTORY_AUDIT` — data tables.
- `STATUS_DASHBOARD` — summary metrics + charts.
- `PRINTABLE_RADIO_LOG` — printable view.

## Form button bindings
I-assign ang script function sa drawing/button per form:

| Form | Search | Submit | Clear | Update |
|---|---|---|---|---|
| AssetMasterForm | `F1_SEARCH` | `F1_SUBMIT` | `F1_CLEAR` | `F1_UPDATE` |
| RadioIssuanceForm | `F2_SEARCH` | `F2_SUBMIT` | `F2_CLEAR` | `F2_UPDATE` |
| MaintenanceEntryForm | `F3_SEARCH` | `F3_SUBMIT` | `F3_CLEAR` | `F3_UPDATE` |
| InventoryAuditForm | `F4_SEARCH` | `F4_SUBMIT` | `F4_CLEAR` | `F4_UPDATE` |
| RadioReturnForm | `F5_SEARCH` | `F5_SUBMIT` | `F5_CLEAR` | `F5_UPDATE` |

## Notes
- Sa `AssetMasterForm`, walang separate search bar — i-type ang Asset Tag sa **B6** field tapos pindutin ang Search.
- Auto-generated ang Asset Tag (`RAD-###`), Transaction ID (`TRX-#####`), at Service ID (`SRV-#####`) on submit.

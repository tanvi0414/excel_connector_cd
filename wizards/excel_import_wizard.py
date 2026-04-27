import base64

from odoo import _, fields, models
from odoo.exceptions import UserError


class ExcelImportWizard(models.TransientModel):
    _name = 'nl.excel.import.wizard'
    _description = 'Excel Import Wizard'

    profile_id = fields.Many2one('nl.excel.profile', string='Profile', required=True, ondelete='cascade')
    file = fields.Binary(string='Excel File', required=True)
    file_name = fields.Char(string='Filename')
    dry_run = fields.Boolean(default=True, help='Validate the file and preview the result without writing to Odoo.')
    result_message = fields.Text(readonly=True)
    error_attachment_id = fields.Many2one('ir.attachment', string='Error File', readonly=True)

    def action_process(self):
        self.ensure_one()
        if not self.file:
            raise UserError(_('Upload an Excel file first.'))
        result = self.profile_id._import_rows(
            workbook_bytes=base64.b64decode(self.file),
            dry_run=self.dry_run,
            filename=self.file_name,
        )
        self.write({
            'result_message': result['message'],
            'error_attachment_id': result['attachment_id'] or False,
        })
        return {
            'type': 'ir.actions.act_window',
            'res_model': self._name,
            'res_id': self.id,
            'view_mode': 'form',
            'target': 'new',
        }

    def action_download_error_file(self):
        self.ensure_one()
        if not self.error_attachment_id:
            raise UserError(_('There is no error file to download.'))
        return {
            'type': 'ir.actions.act_url',
            'url': f'/web/content/{self.error_attachment_id.id}?download=true',
            'target': 'self',
        }

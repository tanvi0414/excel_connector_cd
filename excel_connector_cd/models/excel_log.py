from odoo import fields, models


class ExcelLog(models.Model):
    _name = 'nl.excel.log'
    _description = 'Excel Connector Log'
    _order = 'create_date desc, id desc'

    name = fields.Char(required=True)
    profile_id = fields.Many2one('nl.excel.profile', string='Profile', ondelete='set null', index=True)
    company_id = fields.Many2one('res.company', string='Company', default=lambda self: self.env.company, index=True)
    direction = fields.Selection([
        ('export', 'Export'),
        ('import', 'Import'),
    ], required=True, default='export', index=True)
    state = fields.Selection([
        ('success', 'Success'),
        ('warning', 'Warning'),
        ('error', 'Error'),
    ], required=True, default='success', index=True)
    rows_processed = fields.Integer(default=0)
    rows_failed = fields.Integer(default=0)
    message = fields.Text()
    attachment_id = fields.Many2one('ir.attachment', string='File', ondelete='set null')

from odoo import models, fields, api, _

class AdjustAccountCodesWizard(models.TransientModel):
    _name = 'adjust.account.codes.wizard'
    _description = 'Adjust account codes to 13 digits'

    def action_adjust_codes(self):
        AccountAccount = self.env['account.account']
        accounts = AccountAccount.search([])
        adjusted = 0
        already_ok = 0

        for account in accounts:
            code = account.code or ''
            if len(code) < 13:
                new_code = code.ljust(13, '0')
                account.write({'code': new_code})
                adjusted += 1
            else:
                already_ok += 1

        message = _('Adjusted codes: %s. Already correct: %s.') % (adjusted, already_ok)
        return {
            'type': 'ir.actions.client',
            'tag': 'display_notification',
            'params': {
                'title': _('Account codes adjustment'),
                'message': message,
                'sticky': False,
            }
        }

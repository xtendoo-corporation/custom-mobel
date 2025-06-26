import base64
import xlrd
from odoo import models, fields, api, _

class ImportAccountPlan(models.TransientModel):
    _name = 'import.account.plan'
    _description = 'Importar Plan Contable'

    data_file = fields.Binary(string='Archivo Excel', required=True)
    filename = fields.Char('Nombre del archivo')

    def action_import(self):
        if not self.data_file:
            return

        excel_data = base64.b64decode(self.data_file)
        book = xlrd.open_workbook(file_contents=excel_data)
        sheet = book.sheet_by_index(0)
        AccountAccount = self.env['account.account']

        for row_idx in range(1, sheet.nrows):
            row = sheet.row_values(row_idx)
            cuenta = str(int(row[1])) if row[1] else ''
            subcuenta = str(int(row[2])) if row[2] else ''
            nombre = row[3].strip() if len(row) > 3 else ''
            if not nombre:
                continue
            # Cuenta a 4 dígitos (rellenar a la derecha)
            cuenta = cuenta.ljust(4, '0')
            # Subcuenta a 0 si no hay valor
            subcuenta = subcuenta.zfill(9)
            # Código: concatenar cuenta y subcuenta
            code = cuenta + subcuenta

            # Buscar por código exacto
            existing = AccountAccount.search([('code', '=', code)], limit=1)

            vals = {'code': code}
            if nombre:
                vals['name'] = nombre

            if existing:
                if nombre and existing.name != nombre:
                    existing.write({'name': nombre})
                    print(f"Actualizada cuenta: {code} - {nombre}")
            else:
                AccountAccount.create(vals)
                print(f"Creada cuenta: {code} - {nombre if nombre else '(sin nombre)'}")

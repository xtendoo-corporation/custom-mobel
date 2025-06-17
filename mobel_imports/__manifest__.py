{
    'name': 'Mobel Imports',
    'version': '18.0.1.0.0',
    'category': 'Contacts',
    'summary': 'Importación de clientes y proveedores',
    'author': 'Abraham (Xtendoo)',
    'website': 'https://www.xtendoo.es',
    'license': 'AGPL-3',
    'depends': ['contacts', 'base'],
    'data': [
        'security/ir.model.access.csv',
        'wizards/import_customer_view.xml',
        'wizards/import_supplier_view.xml',
        'views/menu_views.xml',
    ],
    'installable': True,
}

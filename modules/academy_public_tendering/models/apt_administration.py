# -*- coding: utf-8 -*-
###############################################################################
#    License, author and contributors information in:                         #
#    __openerp__.py file at the root folder of this module.                   #
###############################################################################

from openerp import models, fields, api
from openerp.tools.translate import _
from logging import getLogger


_logger = getLogger(__name__)


class AptAdministration(models.Model):
    """ Information about administrations

    Fields:
      name (Char): Human readable name which will identify each record.

    """

    _name = 'apt.administration'
    _description = u'Administration information'

    _rec_name = 'name'
    _order = 'name ASC'

    name = fields.Char(
        string='Name',
        required=True,
        readonly=False,
        index=True,
        default=None,
        help='Enter the name for the administration',
        size=50,
        translate=True
    )

    active = fields.Boolean(
        string='Active',
        required=False,
        readonly=False,
        index=False,
        default=False,
        help=('If the active field is set to false, it will allow you '
              'to hide record without removing it')
    )

    description = fields.Text(
        string='Description',
        required=False,
        readonly=False,
        index=False,
        default=None,
        help='Enter something about the administration',
        translate=True
    )

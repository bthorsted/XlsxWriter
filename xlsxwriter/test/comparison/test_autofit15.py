###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c), 2013-2025, John McNamara, jmcnamara@cpan.org
#

from xlsxwriter.workbook import Workbook

from ..excel_comparison_test import ExcelComparisonTest


class TestCompareXLSXFiles(ExcelComparisonTest):
    """
    Test file created by XlsxWriter against a file created by Excel.

    """

    def setUp(self):

        self.set_filename("autofit15.xlsx")

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename)

        num_format = workbook.add_format({"num_format": "0.000"})

        worksheet = workbook.add_worksheet()

        worksheet.write(0, 0, 123)

        worksheet.write_number(0, 1, 0.123456, num_format)

        worksheet.autofit()

        workbook.close()

        self.assertExcelEqual()

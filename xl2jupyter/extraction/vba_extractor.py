"""VBA module extraction from Excel workbooks."""

from typing import Optional

import xlwings as xw

from xl2jupyter.model.vba import VBAModule
from xl2jupyter.utils.logging import get_logger

logger = get_logger(__name__)


class VBAExtractor:
    """Extract VBA modules from Excel workbooks."""

    def __init__(self, workbook: xw.Book):
        """
        Initialize VBA extractor.

        Args:
            workbook: xlwings Book instance
        """
        self.workbook = workbook

    def extract_all_modules(self) -> list[VBAModule]:
        """
        Extract all VBA modules from the workbook.

        Returns:
            List of VBAModule instances
        """
        modules = []

        try:
            vba_project = self.workbook.api.VBProject
        except Exception as e:
            logger.warning(f"Could not access VBA project: {e}")
            return modules

        try:
            for component in vba_project.VBComponents:
                try:
                    module = self._extract_module(component)
                    if module:
                        modules.append(module)
                except Exception as e:
                    logger.warning(f"Error extracting module {component.Name}: {e}")

        except Exception as e:
            logger.warning(f"Error accessing VBA components: {e}")

        return modules

    def _extract_module(self, component) -> Optional[VBAModule]:
        """
        Extract a single VBA module.

        Args:
            component: VBComponent object

        Returns:
            VBAModule instance or None if extraction fails
        """
        try:
            name = component.Name
            module_type = self._get_module_type(component)
            code = component.CodeModule.Lines(1, component.CodeModule.CountOfLines)

            # Determine sheet name for sheet-specific modules
            sheet_name = None
            if module_type == "Sheet":
                # Try to find the sheet name
                for sheet in self.workbook.sheets:
                    try:
                        if sheet.api.CodeName == name:
                            sheet_name = sheet.name
                            break
                    except Exception:
                        pass

            return VBAModule(
                name=name,
                code=code,
                module_type=module_type,
                sheet_name=sheet_name,
            )

        except Exception as e:
            logger.error(f"Error extracting module {component.Name}: {e}")
            return None

    def _get_module_type(self, component) -> str:
        """
        Get the type of VBA module.

        Args:
            component: VBComponent object

        Returns:
            Module type string
        """
        try:
            # VBComponent.Type constants:
            # 1 = vbext_ct_StdModule (Standard Module)
            # 2 = vbext_ct_ClassModule (Class Module)
            # 3 = vbext_ct_MSForm (UserForm)
            # 100 = vbext_ct_Document (Sheet/Workbook module)

            type_const = component.Type

            if type_const == 1:
                return "Standard"
            elif type_const == 2:
                return "Class"
            elif type_const == 3:
                return "UserForm"
            elif type_const == 100:
                # Check if it's a workbook or sheet module
                name = component.Name.lower()
                if name == "thisworkbook":
                    return "Workbook"
                else:
                    return "Sheet"
            else:
                return "Unknown"

        except Exception:
            return "Unknown"


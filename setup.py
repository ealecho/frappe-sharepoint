from setuptools import setup, find_packages

with open("requirements.txt") as f:
	install_requires = f.read().strip().split("\n")

# get version from __version__ variable in frappe_m365/__init__.py
from frappe_sharepoint import __version__ as version

setup(
	name="frappe_sharepoint",
	version=version,
	description="Universal SharePoint file synchronization for Frappe/ERPNext",
	author="Frappe Community",
	author_email="",
	packages=find_packages(),
	zip_safe=False,
	include_package_data=True,
	install_requires=install_requires
)

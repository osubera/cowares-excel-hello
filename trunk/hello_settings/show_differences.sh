#!/bin/sh

LANG=

diff -s -U 0 wSettingsKeyValue.txt xSettingsKeyValue.txt
diff -s -U 0 wSettingsKeyValueFile.txt xSettingsKeyValueFile.txt
diff -s -U 0 wSettingsKeyValueTable.txt xSettingsKeyValueTable.txt
diff -s -U 0 wSettingsList.txt xSettingsList.txt
diff -s -U 0 wSettingsListFile.txt xSettingsListFile.txt
diff -s -U 0 wSettingsListTable.txt xSettingsListTable.txt
diff -s -U 0 wtestSettingsOnWord.txt xtestSettingsOnExcel.txt

diff -s -U 0 wSettingsKeyValue.txt aSettingsKeyValue.txt
diff -s -U 0 wSettingsKeyValueFile.txt aSettingsKeyValueFile.txt
diff -s -U 0 wSettingsKeyValueTable.txt aSettingsKeyValueTable.txt
diff -s -U 0 wSettingsList.txt aSettingsList.txt
diff -s -U 0 wSettingsListFile.txt aSettingsListFile.txt
diff -s -U 0 wSettingsListTable.txt aSettingsListTable.txt
diff -s -U 0 wtestSettingsOnWord.txt atestSettingsOnAccess.txt

#xSettingsKeyValueSheet.txt
#xSettingsListSheet.txt

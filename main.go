package main

import (
	"encoding/json"
	"errors"
	"fmt"
	"io/fs"
	"log"
	"os"
	"os/exec"
	"path/filepath"
	"strings"
	"syscall"

	"golang.org/x/sys/windows/registry"
)

func createContextMenu(parent string, id string, item *ContextMenu, manifestDir string) (err error) {
	var (
		key     registry.Key
		keyPath = `Software\Classes\Directory\Background\shell` + parent + `\` + id
	)
	if err = deleteRegKeyRecursive(registry.CURRENT_USER, keyPath); err != nil {
		err = fmt.Errorf("failed to delete registry key %q: %w", keyPath, err)
		return
	}
	if key, _, err = registry.CreateKey(registry.CURRENT_USER, keyPath, registry.ALL_ACCESS); err != nil {
		err = fmt.Errorf("failed to create registry key %q: %w", keyPath, err)
		return
	}
	if err = key.SetStringValue("MUIVerb", item.Title); err != nil {
		err = fmt.Errorf("failed to set MUIVerb: %w", err)
		return
	}
	if icon := item.Icon(manifestDir); icon != "" {
		if err = key.SetStringValue("Icon", icon); err != nil {
			err = fmt.Errorf("failed to set Icon: %w", err)
			return
		}
	}
	if item.Extended {
		if err = key.SetStringValue("Extended", ""); err != nil {
			err = fmt.Errorf("failed to set Extended: %w", err)
			return
		}
	}
	if item.Admin {
		if err = key.SetStringValue("HasLUAShield", ""); err != nil {
			err = fmt.Errorf("failed to set HasLUAShield: %w", err)
			return
		}
	}
	if item.Type == ContextMenuType_Folder {
		if err = key.SetStringValue("SubCommands", ""); err != nil {
			err = fmt.Errorf("failed to set SubCommands: %w", err)
			return
		}
		if _, _, err = registry.CreateKey(registry.CURRENT_USER, keyPath+`\shell`, registry.ALL_ACCESS); err != nil {
			err = fmt.Errorf("failed to create registry key %q: %w", keyPath, err)
			return
		}
		for subID, subItem := range item.Items {
			if err = createContextMenu(parent+`\`+id+`\shell`, subID, subItem, manifestDir); err != nil {
				err = fmt.Errorf("failed to create context menu ID %q: %w", subID, err)
				return
			}
		}
	} else {
		keyPath += `\command`
		if err = deleteRegKeyRecursive(registry.CURRENT_USER, keyPath); err != nil {
			err = fmt.Errorf("failed to delete registry key %q: %w", keyPath, err)
			return
		}
		if key, _, err = registry.CreateKey(registry.CURRENT_USER, keyPath, registry.ALL_ACCESS); err != nil {
			err = fmt.Errorf("failed to create registry key %q: %w", keyPath, err)
			return
		}
		if err = key.SetExpandStringValue("", item.CommandString(manifestDir)); err != nil {
			err = fmt.Errorf("failed to set command string: %w", err)
			return
		}
	}
	return
}

func run() (err error) {
	var (
		manifestPath, manifestDir string
		manifestData              []byte
		manifest                  Manifest
	)
	if manifestPath, err = findManifest(); err != nil {
		return
	}
	manifestDir = filepath.Dir(manifestPath)
	if manifestData, err = os.ReadFile(manifestPath); err != nil {
		err = fmt.Errorf("failed to read manifest.json: %w", err)
		return
	}
	if err = json.Unmarshal(manifestData, &manifest); err != nil {
		err = fmt.Errorf("failed to parse manifest.json: %w", err)
		return
	}
	for id, item := range manifest.Items {
		if err = createContextMenu("", id, item, manifestDir); err != nil {
			err = fmt.Errorf("failed to create context menu ID %q: %w", id, err)
			return
		}
	}
	return
}

func findManifest() (manifestPath string, err error) {
	const manifestFilename = "manifest.json"
	var (
		fi   fs.FileInfo
		fp   string
		terr error
	)
	if fp, terr = os.Getwd(); terr == nil {
		manifestPath = filepath.Join(fp, manifestFilename)
		if fi, terr = os.Stat(manifestPath); terr == nil && !fi.IsDir() {
			return
		}
	}
	if fp, terr = os.Executable(); terr == nil {
		manifestPath = filepath.Join(filepath.Dir(fp), manifestFilename)
		if fi, terr = os.Stat(manifestPath); terr == nil && !fi.IsDir() {
			return
		}
	}
	err = fmt.Errorf("manifest.json not found: %w", os.ErrNotExist)
	return
}

var _nircmdPath string

func findNircmd() (nircmdPath string, err error) {
	const nircmdFilename = "nircmd.exe"
	var (
		fi   fs.FileInfo
		fp   string
		terr error
	)
	if _nircmdPath != "" {
		nircmdPath = _nircmdPath
		return
	}
	defer func() {
		_nircmdPath = nircmdPath
	}()
	if fp, terr = os.Getwd(); terr == nil {
		nircmdPath = filepath.Join(fp, nircmdFilename)
		if fi, terr = os.Stat(nircmdPath); terr == nil && !fi.IsDir() {
			return
		}
		nircmdPath = filepath.Join(fp, "bin", nircmdFilename)
		if fi, terr = os.Stat(nircmdPath); terr == nil && !fi.IsDir() {
			return
		}
	}
	if fp, terr = os.Executable(); terr == nil {
		nircmdPath = filepath.Join(filepath.Dir(fp), nircmdFilename)
		if fi, terr = os.Stat(nircmdPath); terr == nil && !fi.IsDir() {
			return
		}
		nircmdPath = filepath.Join(filepath.Dir(fp), "bin", nircmdFilename)
		if fi, terr = os.Stat(nircmdPath); terr == nil && !fi.IsDir() {
			return
		}
	}
	if nircmdPath, terr = exec.LookPath(nircmdFilename); terr == nil {
		return
	}
	err = fmt.Errorf("nircmd.exe not found: %w", os.ErrNotExist)
	return
}

func main() {
	if err := run(); err != nil {
		log.Fatal(err)
	}
}

type ContextMenu struct {
	Type      ContextMenuType         `json:"type"`
	Title     string                  `json:"title"`
	IconPath  string                  `json:"iconPath"`
	IconIndex *int                    `json:"iconIndex,omitempty"`
	Extended  bool                    `json:"extended"`
	Admin     bool                    `json:"admin"`
	Command   []string                `json:"command,omitempty"`
	Items     map[string]*ContextMenu `json:"items,omitempty"`
}

type ContextMenuType string

const (
	ContextMenuType_Item   ContextMenuType = "item"
	ContextMenuType_Folder ContextMenuType = "folder"
)

type Manifest struct {
	Items map[string]*ContextMenu `json:"items"`
}

func (c ContextMenu) Icon(manifestDir string) string {
	iconPath := c.IconPath
	if iconPath == "" {
		return ""
	}
	iconPath = strings.ReplaceAll(iconPath, "${manifestFolder}", manifestDir)
	iconPath = quoteWindowsPath(iconPath)
	if c.IconIndex != nil {
		iconPath = fmt.Sprintf("%s,%d", iconPath, *c.IconIndex)
	}
	return iconPath
}

func (c ContextMenu) CommandString(manifestDir string) string {
	var (
		err        error
		nircmdPath string
		command    []string
	)
	if c.Admin {
		if nircmdPath, err = findNircmd(); err != nil {
			log.Fatal(err)
		} else {
			command = append(command, quoteWindowsPath(nircmdPath), "elevate")
		}
	}
	for _, part := range c.Command {
		part = strings.ReplaceAll(part, "${manifestFolder}", manifestDir)
		if strings.ContainsAny(part, " %") {
			part = quoteWindowsPath(part)
		}
		command = append(command, part)
	}
	return strings.Join(command, " ")
}

func deleteRegKeyRecursive(k registry.Key, path string) (err error) {
	var (
		key, emptyKey registry.Key
		subKeyNames   []string
	)
	if key, err = registry.OpenKey(k, path, registry.ALL_ACCESS); err != nil {
		if errors.Is(err, syscall.ENOENT) {
			err = nil
			return
		}
		err = fmt.Errorf("deleteRegKeyRecursive failed to open key path %q: %w", path, err)
		return
	}
	defer func() {
		if key != emptyKey {
			key.Close()
		}
	}()
	if subKeyNames, err = key.ReadSubKeyNames(0); err != nil {
		err = fmt.Errorf("deleteRegKeyRecursive failed to get subkeys of path %q: %w", path, err)
		return
	}
	for _, subKeyName := range subKeyNames {
		if err = deleteRegKeyRecursive(key, subKeyName); err != nil {
			err = fmt.Errorf("deleteRegKeyRecursive failed to delete subkey %q of path %q: %w", subKeyName, path, err)
			return
		}
	}
	key.Close()
	key = emptyKey
	if err = registry.DeleteKey(k, path); err != nil {
		if errors.Is(err, syscall.ENOENT) {
			err = nil
			return
		}
		err = fmt.Errorf("deleteRegKeyRecursive failed to delete key path %q: %w", path, err)
		return
	}
	return
}

func quoteWindowsPath(path string) string {
	return `"` + path + `"`
}

package main

import (
	"fmt"
	"log"
	"strings"
	"syscall"
	"time"
	"unsafe"

	"github.com/go-ole/go-ole"
	"github.com/moutend/go-wca/pkg/wca"
)

func main() {
	// Disable console window selection to prevent hanging.
	if err := disableConsoleQuickEdit(); err != nil {
		log.Printf("Warning: Failed to disable console quick edit mode: %v", err)
	}

	// Initialize COM library.
	if err := ole.CoInitialize(0); err != nil {
		panic(fmt.Errorf("failed to initialize COM library: %w", err))
	}
	defer ole.CoUninitialize()

	// Update volume of Rocksmith device every 5 seconds
	for {
		err := setRocksmithDeviceVolume()
		if err != nil {
			log.Fatalf("Error listing microphones: %v", err)
		}
		// Sleep for 5 seconds
		<-time.After(5 * time.Second)
	}

}

func setRocksmithDeviceVolume() error {
	// Create the MMDeviceEnumerator to access audio devices.
	var enumerator *wca.IMMDeviceEnumerator
	if err := wca.CoCreateInstance(
		wca.CLSID_MMDeviceEnumerator,
		0,
		wca.CLSCTX_ALL,
		wca.IID_IMMDeviceEnumerator,
		&enumerator,
	); err != nil {
		return fmt.Errorf("failed to create MMDeviceEnumerator: %w", err)
	}
	defer enumerator.Release()

	// Get the collection of audio capture devices (microphones).
	var collection *wca.IMMDeviceCollection
	if err := enumerator.EnumAudioEndpoints(wca.ECapture, wca.DEVICE_STATE_ACTIVE, &collection); err != nil {
		return fmt.Errorf("failed to enumerate audio endpoints: %w", err)
	}
	defer collection.Release()

	// Get the count of audio capture devices.
	var count uint32
	if err := collection.GetCount(&count); err != nil {
		return fmt.Errorf("failed to get device count: %w", err)
	}

	for i := 0; i < int(count); i++ {
		// Select device
		var device *wca.IMMDevice
		if err := collection.Item(uint32(i), &device); err != nil {
			return fmt.Errorf("failed to get device at index %d: %w", i, err)
		}
		defer device.Release()

		// // Get the device ID.
		// var deviceId string
		// if err := device.GetId(&deviceId); err != nil {
		// 	return nil, fmt.Errorf("failed to get device ID: %w", err)
		// }
		// fmt.Printf("Device ID: %s\n", deviceId)

		// Get the property store of the device to retrieve friendly name.
		var propertyStore *wca.IPropertyStore
		if err := device.OpenPropertyStore(wca.STGM_READ, &propertyStore); err != nil {
			return fmt.Errorf("failed to open property store: %w", err)
		}
		defer propertyStore.Release()

		// Get the friendly name of the device.
		var propValue wca.PROPVARIANT
		if err := propertyStore.GetValue(&wca.PKEY_Device_FriendlyName, &propValue); err != nil {
			return fmt.Errorf("failed to get friendly name: %w", err)
		}

		// Add to result if it's a Rocksmith device
		deviceName := propValue.String()
		if strings.Contains(deviceName, "Rocksmith") {
			// Set volume of Rocksmith device
			var audioEndpoint *wca.IAudioEndpointVolume
			if err := device.Activate(wca.IID_IAudioEndpointVolume, wca.CLSCTX_ALL, nil, &audioEndpoint); err != nil {
				return fmt.Errorf("failed to activate audio endpoint: %w", err)
			}
			defer audioEndpoint.Release()

			if err := audioEndpoint.SetMasterVolumeLevelScalar(1.0, nil); err != nil {
				return fmt.Errorf("failed to set volume to max: %w", err)
			}
			fmt.Printf("Set volume to max for device: %s\n", deviceName)
		}
	}

	return nil
}

func disableConsoleQuickEdit() error {
	k32 := syscall.NewLazyDLL("kernel32.dll")
	setConsoleMode := k32.NewProc("SetConsoleMode")
	getConsoleMode := k32.NewProc("GetConsoleMode")
	stdinHandle := syscall.Handle(syscall.Stdin)

	var mode uint32
	ret, _, err := getConsoleMode.Call(uintptr(stdinHandle), uintptr(unsafe.Pointer(&mode)))
	if ret == 0 {
		return err
	}

	// Disable quick edit mode to prevent console hanging.
	const ENABLE_QUICK_EDIT_MODE = 0x0040
	mode &^= ENABLE_QUICK_EDIT_MODE

	ret, _, err = setConsoleMode.Call(uintptr(stdinHandle), uintptr(mode))
	if ret == 0 {
		return err
	}

	return nil
}

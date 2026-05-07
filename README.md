# WOL Sender
A lightweight 100KB Wake-on-LAN (WOL) utility that supports Win95 to Win11, written in pure C using the native Win32 API.
* Allows you to manage a list of computers, check their online status via ICMP or TCP, and wake them up using Magic Packets.

<img src="https://github.com/user-attachments/assets/3f5c3a20-b93f-4f3c-a8d4-c6b025ef43a3" width="30%"></img>
<img src="https://github.com/user-attachments/assets/677c5b2f-ab75-4999-ae5a-7159d9d46f98" width="30%"></img> 
<img src="https://github.com/user-attachments/assets/fd4703bc-d0e6-45a7-a9ea-879d746ca606" width="30%"></img>

## Features
* **Native & Fast:** <em>Built with pure Win32 API and Winsock. No heavy frameworks, no .NET requirements, and a microscopic binary size.</em>
      
*  **Status Checking:** <em>Built-in reachability testing using ICMP (Ping) and TCP Port 445 (SMB) to verify if a machine is already awake.</em>
      
*  **Contact Management:** <em>Save, edit, and remove hosts. Data is persisted in a portable wol_contacts.ini file.</em>
      
*  **Batch Actions:** <em>Select multiple entries to wake or check status simultaneously.</em>
      
*  **Import/Export:** <em>Easily move your contact list between machines using the built-in INI import/export functions.</em>
      
*  **Tiny Footprint:** <em>Highly optimized for size, targeting compatibility back to older i386 architectures. 100KB size.</em>

## Compile & Build

The project is designed to be compiled with MinGW/GCC.
The following flags are used to strip unnecessary bloat and optimize for the smallest possible executable size.
1.  **Prerequisites:**
    * [MinGW](https://sourceforge.net/projects/mingw/) and [MSYS2](https://www.msys2.org/) installed.

2.  **Compile**
    * To compile the source (assuming you have a `resources.o` for the application icon):
    ```bash
    gcc wol_sender.c resources.o -o wolsender.exe -mwindows -Os -s -march=i386 -mno-sse -mno-sse2 -ffunction-sections -fdata-sections "-Wl,--gc-sections" -lwsock32 -lcomdlg32 -lcomctl32 -lshell32
    ```
3.  **Compress (Optional)**
    * To further shrink the binary, use UPX with LZMA compression:
    ```bash
    upx.exe --best --lzma wolsender.exe
    ```

##  Requirements
* **OS:** Windows (95 or newer)
* **Network:** Wake-On-Lan (Magic Packets) must be enabled in the target computer's BIOS/UEFI and Network Adapter settings.

## License
<em>This project is open-source. Feel free to fork, modify, and use it for your own networking needs.
Contributions, issues, and feature requests are welcome! Feel free to check the [issues page](https://github.com/Treblewolf/wol-sender/issues).


Distributed under the MIT License. See `LICENSE` file for more information.
</em>

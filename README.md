![Logo](https://www.audiostation.org/images/logo.png)

# Audiostation

Audiostation is an old school media player that can be installed on Windows systems. The program can play all common audio files like (*.mp3, *.wav, etc). It has a record function and has all the elements to play all the music you want to play.

#### Contribute to the project

Your contributions are greatly welcome

- [Translate Audiostation](/docs/translations.md)
- [Contribute fixes](/docs/contribution.md)
- [Report errors and suggest improvements](/docs/report-bugs.md)
- [Donate](/docs/donations.md)

#### Application File Support

**Supported Audio Files**

| Type  | Description                          | Convertion ** |
| ----- | ------------------------------------ | ------------- |
| MP2   | MPEG-1 Layer 2                       |               |
| M4A   | MPEG-4 Layer 4 Audio                 |               |
| RA    | Real Audio                           | Needed        |
| RM    | Real Media                           | Needed        |
| CDA   | CD Audio                             |               |
| WAV   | Microsoft WaveForm Audio             |               |
| AIF   | Audio Interchange File               |               |
| AAC   | Advanced Audio Coding                |               |
| SND   | Sun Microsystems Sound               |               |
| AU    | Sun Microsystems Audio               |               |
| WMA   | Windows Media Audio                  |               |
| RMI   | Musical Instrument Digital Interface | Needed        |
| ACT   | Voice File Format                    | Needed        |
| CAF   | Apple Core Format                    | Needed        |
| WSAUD | Westwood Studios Audio               | Needed        |
| W64   | Sony Wave64                          | Needed        |
| OGG   | OGG Sound File                       | Needed        |
| AMO   | Sony Opend Audio                     | Needed        |
| VOC   | Creative Voice                       | Needed        |
| MID   | Midi                                 |               |
| KAR   | Karaoke Midi File                    |               |
| SID   | Commodore 64 Sound Files             |               |
| MUS   | Sibra-Soft Beep Symphony             |               |

**Supported Playlist Files**

| Type | Description                        | Convertion ** |
| ---- | ---------------------------------- | ------------- |
| PLS  | ShoutCast Playlist File            |               |
| WPL  | Windows Media Player Playlist File |               |
| APL  | Audiostation Playlist File         |               |
| M3U  | Playlist File                      |               |

** De convertion column tells you if a file must be converted before it can be played by Audiostation.

#### Installation & Packages

The Audiostation installation package is available at the following recources

- Website package
- Sourceforge package
- Softpedia package
- Chocolatey package

You can choose one of these packages to download and install Audiostation.  
Approximate package size: **160 MB**

Use our: [Installation Guide](/docs/how-to-install.md) to see how to install Audiostation

#### Development

The Audiostation application is writen in **Microsoft Visual Basic 6** and therefore requires various dependencies, you must use the dependency installer from our Github to make sure you use the correct dependencies.

If you get a error when opening the Visual Basic project you must reinstall the dependencies. 

Below you will find a list of the most common used dependencies

| Filename          | Description                                                 | Type |
| ----------------- | ----------------------------------------------------------- | ---- |
| audio_sniffer     | Ads the virutal audio recorder to 32-bit computers.         | dll  |
| audio_sniffer-x64 | Ads the virtual audio recorder to 64-bit computers          | dll  |
| basswavapi        | Core library for capturing the wav output of the soundcard. | dll  |
| bass              | Core library for capturing the wav output of the soundcard. | dll  |
| LaVolpeAlphaImg2  | Picture control for showing Alpha images like PNG, etc.     | ocx  |
| IsAnalogLibrary   | Holds the controls for showing digital displays, etc.       | ocx  |
| MbPrgBar          | Progressbar control                                         | ocx  |
| midifl32          | Core midi file library                                      | ocx  |
| midiio32          | Core midi input/output library                              | ocx  |
| mscomctl          | Microsoft common controls library - SP5                     | ocx  |
| win32             | Holds all the icons used for Audiostation                   | dll  |
| unzip32           | Library for unzipping zip files                             | dll  |

#### Installation Package Contents

The Audiostation installation package contains the following files:

**Contents**

- Audio capture (2,77 MB)
- Midi Soundfont (141 MB)
- File Assosiaction Icon Files (12 KB)
- Language Files (3 KB)
- Midi, Wav and Mus Sample Files (5 MB)
- Third Party Support Files 22,5 MB)
- Application Dependencies (20,5 MB)
- Main Application File (2,16 MB)
- Application Updater Program (413 KB)

Total package size on disk: 195 MB  
Total package size compressed: 160 MB
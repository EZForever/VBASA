# VBASA

*Standalone Visual Basic for Applications (VBA) 7.1 IDE with EXE build support*

> [!NOTE]
> This repository is part of my *Year-end Things Clearance 2024* code dump, which implies the following things:
> 
> 1. The code in this repo is from one of my personal hobby projects (mainly) developed in 2024
> 2. Since it's a hobby project, expect PoC / alpha-quality code, and an *astonishing* lack of documentation
> 3. The project is archived by the time this repo goes public; there will (most likely) be no future updates
> 
> In short, if you wish to use the code here for anything good, think twice.

Yes, you read that right; a *standalone* VBA IDE that works on *standalone* VBA project files and (with proper license) builds *standalone* EXEs, all without involving any host program. It even supports x64!

This project, finished in May 2024, marks an end to my years-long research around VBA and the possibility of using it in modern days as a general-purpose programming language. There is so, *so* much to document (likely a hundred of pages of docs), and writing all that down during an *Year-end Things Clearance* is just not possible, so here's a repository with some publish-able code that you can toy with first. Enjoy.

*TODO: Questions to be answered in this document:*

- "Why VBA? Isn't that some 30-year old technology? Don't you have better things to do?"
- "Isn't VBA SDK from Microsoft already support this?"
    - Well, I could answer that one right now: for one, that SDK gets discontinued after releasing 32-bit VBA6; more importantly, COM registrations are a PITA to deal with
- How to actually build & use this IDE (and where you could get VBA redist packages; yes, those exist)
- Explain why you need a cryptic license key for EXE build to work (spoiler: Microsoft lawyers would be pretty mad otherwise)
- Document actual research & findings involved to make this thing work as it is (e.g. how Access MDE / Source Code Removal works)
- Document *other* research & findings that don't make it into this repository (e.g. how different is VBA7 P-Code engine compared to VB6/VBA6)
- ~~When will *Project #5432* be released~~


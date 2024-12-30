#include "vbe.h"

#pragma warning(push)
#pragma warning(disable: 4192) // Duplicate of system types in VBE.TLB

// Microsoft Vbe UI 7.1 Object Library
#import "VBEUI.TLB" \
    implementation_only auto_rename

// Microsoft Visual Basic for Applications Extensibility 5.3
#import "VBE6EXT.OLB" \
    implementation_only

// UNOFFICIAL Microsoft Visual Basic for Applications 6.0+ Type Library
#import "VBE.TLB" \
    implementation_only auto_rename

#pragma warning(pop)


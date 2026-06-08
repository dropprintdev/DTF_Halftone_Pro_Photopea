var globalSettings;

function doRunDTFPrep() {
    try {
        if (app.documents.length === 0) {
            app.echoToOE("ERROR: Please open an image first!");
            return;
        }
        app.activeDocument.suspendHistory("DTF Halftone Prep", "doRunDTFPrepCore()");
    } catch(e) {
        app.echoToOE("ERROR: " + e.toString());
    }
}

function doRunDTFPrepCore() {
    try {
        var settings = globalSettings;
        var doc = app.activeDocument;
        
        if (doc.activeLayer.isBackgroundLayer) {
            doc.activeLayer.isBackgroundLayer = false;
        }
        
        try {
            if (doc.activeLayer.transparentPixelsLocked) {
                doc.activeLayer.transparentPixelsLocked = false;
            }
        } catch(e) {}
        
        applyLevels(settings.blackPoint, settings.grayPoint, settings.whitePoint, settings.outBlack, settings.outWhite);
        
        if (settings.shadowBoost > 0) {
            applyShadowBoost(settings.shadowBoost);
        }
        
        if (settings.enableKnockout) {
            applyColorKnockout(settings.knockoutR, settings.knockoutG, settings.knockoutB);
        }
        
        if (settings.enableHalftone) {
            applyHalftone(settings.halftoneShape, settings.frequency, settings.angle);
        }
        
        if (settings.defringe) {
            applyDefringe();
        }
    } catch(e) {
        app.echoToOE("ERROR during processing: " + e.toString());
    }
}

function applyLevels(bp, gp, wp, obp, owp) {
    try {
        var desc = new ActionDescriptor();
        var ref = new ActionReference();
        ref.putEnumerated(charIDToTypeID("Chnl"), charIDToTypeID("Chnl"), charIDToTypeID("Cmps"));
        var descAdjustment = new ActionDescriptor();
        descAdjustment.putReference(charIDToTypeID("Chnl"), ref);
        
        var listInput = new ActionList();
        listInput.putInteger(bp);
        listInput.putInteger(wp);
        descAdjustment.putList(charIDToTypeID("Inpt"), listInput);
        descAdjustment.putDouble(charIDToTypeID("Gmm "), gp);
        
        var listOutput = new ActionList();
        listOutput.putInteger(obp);
        listOutput.putInteger(owp);
        descAdjustment.putList(charIDToTypeID("Otpt"), listOutput);
        
        var listLevels = new ActionList();
        listLevels.putObject(charIDToTypeID("LvlA"), descAdjustment);
        desc.putList(charIDToTypeID("Adjs"), listLevels);
        
        executeAction(charIDToTypeID("Lvls"), desc, DialogModes.NO);
    } catch(e) {
        // Fallback to DOM
        app.activeDocument.activeLayer.adjustLevels(bp, wp, gp, obp, owp);
    }
}

function applyShadowBoost(amount) {
    try {
        var desc = new ActionDescriptor();
        var shadowDesc = new ActionDescriptor();
        shadowDesc.putUnitDouble(charIDToTypeID("Amnt"), charIDToTypeID("#Prc"), amount);
        shadowDesc.putUnitDouble(charIDToTypeID("Wdth"), charIDToTypeID("#Prc"), 50);
        shadowDesc.putInteger(charIDToTypeID("Rds "), 30);
        desc.putObject(charIDToTypeID("Shdw"), charIDToTypeID("Adpt"), shadowDesc);
        executeAction(stringIDToTypeID("adaptCorrect"), desc, DialogModes.NO);
    } catch(e) {}
}

function doRunDTFPrepAfterKnockout() {
    try {
        var doc = app.activeDocument;
        doc.activeLayer.rasterize(RasterizeType.ENTIRELAYER);
        try {
            var oldLayer = doc.artLayers.getByName("DTF_OLD_LAYER");
            oldLayer.remove();
        } catch(e) {}
        
        doRunDTFPrepCore();
    } catch(e) {
        app.echoToOE("ERROR in Knockout cleanup: " + e.toString());
    }
}

function applyHalftone(shape, freq, angle) {
    var doc = app.activeDocument;
    var originalLayer = doc.activeLayer;

    var tempDoc = doc.duplicate("TempHalftoneDoc");
    app.activeDocument = tempDoc;
    var tempLayer = tempDoc.activeLayer;
    
    if (tempLayer.isBackgroundLayer) tempLayer.isBackgroundLayer = false;

    tempLayer.transparentPixelsLocked = true;
    var blackColor = new SolidColor();
    blackColor.rgb.red = 0; blackColor.rgb.green = 0; blackColor.rgb.blue = 0;
    tempDoc.selection.selectAll();
    tempDoc.selection.fill(blackColor);
    tempDoc.selection.deselect();
    tempLayer.transparentPixelsLocked = false;

    var bgLayer = tempDoc.artLayers.add();
    var whiteColor = new SolidColor();
    whiteColor.rgb.red = 255; whiteColor.rgb.green = 255; whiteColor.rgb.blue = 255;
    tempDoc.selection.selectAll();
    tempDoc.selection.fill(whiteColor);
    tempDoc.selection.deselect();
    bgLayer.move(tempLayer, ElementPlacement.PLACEAFTER);

    tempDoc.flatten();
    
    // Photopea Halftone Method
    // We try Bitmap conversion, if it fails, we fall back to Photopea's Color Halftone filter
    try {
        tempDoc.changeMode(ChangeMode.GRAYSCALE);
        var options = new BitmapConversionOptions();
        options.resolution = tempDoc.resolution;
        options.method = BitmapConversionType.HALFTONESCREEN;
        options.angle = angle;
        options.frequency = freq;
        if (shape === "Round") options.shape = BitmapHalfToneType.ROUND;
        else options.shape = BitmapHalfToneType.ROUND;
        
        tempDoc.changeMode(ChangeMode.BITMAP, options);
        tempDoc.changeMode(ChangeMode.GRAYSCALE);
    } catch(e) {
        // Fallback: Color Halftone (Photopea might support this better)
        var desc = new ActionDescriptor();
        desc.putInteger(charIDToTypeID("Rds "), 4); // Dot size based on freq
        desc.putInteger(charIDToTypeID("Ang1"), angle);
        desc.putInteger(charIDToTypeID("Ang2"), angle);
        desc.putInteger(charIDToTypeID("Ang3"), angle);
        desc.putInteger(charIDToTypeID("Ang4"), angle);
        executeAction(charIDToTypeID("ClrH"), desc, DialogModes.NO);
    }

    tempDoc.activeLayer.invert();
    tempDoc.selection.selectAll();
    tempDoc.selection.copy();
    tempDoc.close(SaveOptions.DONOTSAVECHANGES);

    app.activeDocument = doc;
    doc.activeLayer = originalLayer;

    addLayerMask();
    pasteIntoLayerMask();
}

function addLayerMask() {
    try {
        var desc = new ActionDescriptor();
        desc.putClass(charIDToTypeID("Nw  "), charIDToTypeID("Chnl"));
        var ref = new ActionReference();
        ref.putEnumerated(charIDToTypeID("Chnl"), charIDToTypeID("Chnl"), charIDToTypeID("Msk "));
        desc.putReference(charIDToTypeID("At  "), ref);
        desc.putEnumerated(charIDToTypeID("Usng"), charIDToTypeID("UsrM"), charIDToTypeID("RvlA"));
        executeAction(charIDToTypeID("Mk  "), desc, DialogModes.NO);
    } catch(e) {}
}

function pasteIntoLayerMask() {
    try {
        var desc = new ActionDescriptor();
        var ref = new ActionReference();
        ref.putEnumerated(charIDToTypeID("Chnl"), charIDToTypeID("Chnl"), charIDToTypeID("Msk "));
        desc.putReference(charIDToTypeID("null"), ref);
        executeAction(charIDToTypeID("slct"), desc, DialogModes.NO);

        app.activeDocument.paste();

        var desc2 = new ActionDescriptor();
        var ref2 = new ActionReference();
        ref2.putEnumerated(charIDToTypeID("Chnl"), charIDToTypeID("Chnl"), charIDToTypeID("RGB "));
        desc2.putReference(charIDToTypeID("null"), ref2);
        executeAction(charIDToTypeID("slct"), desc2, DialogModes.NO);
        
        var desc3 = new ActionDescriptor();
        var ref3 = new ActionReference();
        ref3.putEnumerated(charIDToTypeID("Chnl"), charIDToTypeID("Ordn"), charIDToTypeID("Trgt"));
        desc3.putReference(charIDToTypeID("null"), ref3);
        desc3.putBoolean(charIDToTypeID("Aply"), true);
        executeAction(charIDToTypeID("Dlt "), desc3, DialogModes.NO);
    } catch(e) {}
}

function applyDefringe() {
    try {
        var desc = new ActionDescriptor();
        var ref = new ActionReference();
        ref.putProperty(charIDToTypeID("Chnl"), charIDToTypeID("fsel"));
        desc.putReference(charIDToTypeID("null"), ref);
        var ref2 = new ActionReference();
        ref2.putEnumerated(charIDToTypeID("Chnl"), charIDToTypeID("Chnl"), charIDToTypeID("Trsp"));
        desc.putReference(charIDToTypeID("T   "), ref2);
        executeAction(charIDToTypeID("setd"), desc, DialogModes.NO);
        
        app.activeDocument.selection.contract(1);
        app.activeDocument.selection.invert();
        app.activeDocument.selection.clear();
        app.activeDocument.selection.deselect();
    } catch(e) {}
}

"""Application-owned interface strings; document contents are never translated."""
from .model import ValidationError

HINDI = {
    "Show DevLys preview…": "DevLys पूर्वावलोकन दिखाएँ…",
    "DevLys preview": "DevLys पूर्वावलोकन",
    "Live DevLys preview": "DevLys का सीधा पूर्वावलोकन",
    "Preview: {field}": "पूर्वावलोकन: {field}",
    "Select a text field": "कोई पाठ फ़ील्ड चुनें",
    "Keyboard codes": "कीबोर्ड कोड",
    "Hindi display (DevLys 010)": "हिन्दी रूप (DevLys 010)",
    "Document display (literal text)": "दस्तावेज़ का रूप (अक्षर और अंक)",
    "The same keyboard codes are displayed using the document font. No Unicode conversion is performed.": "उन्हीं कीबोर्ड कोड को दस्तावेज़ के फ़ॉन्ट में दिखाया जाता है। यूनिकोड रूपांतरण नहीं किया जाता।",
    "DevLys 010 is not installed. Install it with Font Book for an accurate Hindi preview.": "DevLys 010 स्थापित नहीं है। सही हिन्दी पूर्वावलोकन के लिए इसे Font Book से स्थापित करें।",
    "Type DevLys keyboard codes. The shared preview follows the selected field.": "DevLys के कीबोर्ड कोड लिखें। साझा पूर्वावलोकन चुने हुए फ़ील्ड का पाठ दिखाता है।",
    "Documents are saved in output/ as clientname_data.docx.": "दस्तावेज़ output/ में clientname_data.docx नाम से सहेजे जाते हैं।",
    "This desktop form accepts DevLys text only. Unicode client files cannot be converted automatically.": "यह फ़ॉर्म केवल DevLys पाठ स्वीकार करता है। यूनिकोड ग्राहक फ़ाइलों का स्वतः रूपांतरण नहीं किया जा सकता।",
    "Land record documents": "भूमि अभिलेख दस्तावेज़",
    "Interface language": "इंटरफ़ेस भाषा",
    "New": "नया", "Open JSON…": "JSON खोलें…", "Save JSON…": "JSON सहेजें…", "Load sample": "नमूना लोड करें",
    "Client details": "ग्राहक का विवरण", "Client name": "ग्राहक का नाम", "Father's name": "पिता का नाम",
    "Age (years)": "आयु (वर्ष)", "Age": "आयु", "Caste": "जाति", "Address": "पता", "Village": "ग्राम",
    "Tehsil": "तहसील", "District": "ज़िला", "Document date (YYYY-MM-DD)": "दस्तावेज़ दिनांक (YYYY-MM-DD)",
    "Bahi number (one per document)": "बही क्रमांक (पूरे दस्तावेज़ के लिए एक)", "Bahi number": "बही क्रमांक",
    "Text and fonts": "पाठ और फ़ॉन्ट", "Document input font": "दस्तावेज़ इनपुट फ़ॉन्ट",
    "DevLys mode requires already encoded keyboard text, not Unicode Hindi. No automatic conversion is performed.": "DevLys के लिए पहले से एन्कोड किए कीबोर्ड कोड दें, यूनिकोड हिन्दी नहीं। स्वतः रूपांतरण नहीं किया जाता।",
    "Changing this selection does not convert text already entered.": "यह चयन बदलने से पहले लिखा पाठ परिवर्तित नहीं होता।",
    "Land records": "भूमि अभिलेख", "Add…": "जोड़ें…", "Edit…": "संपादित करें…", "Remove": "हटाएँ",
    "Double-click a row to edit": "संपादन के लिए पंक्ति पर दो बार क्लिक करें",
    "Khasra number": "खसरा क्रमांक", "Area (ha)": "क्षेत्रफल (हेक्टेयर)", "Area (hectares)": "क्षेत्रफल (हेक्टेयर)", "Area": "क्षेत्रफल", "Lagan": "लगान",
    "{count} records  •  Total area: {area} ha  •  Total Lagan: {lagan}": "{count} अभिलेख  •  कुल क्षेत्रफल: {area} हेक्टेयर  •  कुल लगान: {lagan}",
    "Generate new DOCX…": "नया DOCX बनाएँ…",
    "Uses the original template in templates/. Review fixed legal declarations and valuations before sharing.": "templates/ के मूल टेम्पलेट का उपयोग होता है। साझा करने से पहले स्थिर कानूनी घोषणाओं और मूल्यांकन की समीक्षा करें।",
    "Enter client details or load sample data.": "ग्राहक का विवरण भरें या नमूना डेटा लोड करें।", "Untitled": "बिना शीर्षक",
    "Add land record": "भूमि अभिलेख जोड़ें", "Edit land record": "भूमि अभिलेख संपादित करें",
    "Identifiers keep leading zeros and slashes.\nUse a dot for decimals; no commas.": "क्रमांकों में शुरुआती शून्य और स्लैश सुरक्षित रहेंगे।\nदशमलव के लिए बिंदु लगाएँ; अल्पविराम नहीं।",
    "Cancel": "रद्द करें", "Apply": "लागू करें", "Check land record": "भूमि अभिलेख जाँचें", "Check client data": "ग्राहक डेटा जाँचें",
    "Select one land record to edit.": "संपादन के लिए एक भूमि अभिलेख चुनें।",
    "Select one or more land records to remove.": "हटाने के लिए एक या अधिक भूमि अभिलेख चुनें।",
    "Unsaved client data": "ग्राहक डेटा सहेजा नहीं गया", "Save your changes before continuing?": "आगे बढ़ने से पहले बदलाव सहेजें?",
    "New client.": "नया ग्राहक।", "Open client JSON": "ग्राहक JSON खोलें", "Client JSON": "ग्राहक JSON",
    "Cannot open client data": "ग्राहक डेटा नहीं खुल सका", "Sample loaded. Save a copy for your client.": "नमूना लोड हुआ। अपने ग्राहक के लिए उसकी प्रति सहेजें।",
    "Loaded {path}": "लोड हुआ: {path}", "Save client JSON": "ग्राहक JSON सहेजें", "Cannot save client data": "ग्राहक डेटा नहीं सहेजा जा सका", "Saved {path}": "सहेजा गया: {path}",
    "Generate a new Word document": "नया Word दस्तावेज़ बनाएँ", "Word document": "Word दस्तावेज़", "Cannot generate document": "दस्तावेज़ नहीं बन सका",
    "Generated {name} — updated all {count} land tables.": "{name} बनाया गया — सभी {count} भूमि तालिकाएँ अद्यतन हुईं।",
    "Document generated": "दस्तावेज़ बन गया", "Saved {path}\n\nAll three land tables were updated, including the nested table.": "सहेजा गया: {path}\n\nभीतरी तालिका सहित तीनों भूमि तालिकाएँ अद्यतन हुईं।",
    "Older client loaded. Enter age and caste before saving or generating.": "पुराना ग्राहक डेटा लोड हुआ। सहेजने या दस्तावेज़ बनाने से पहले आयु और जाति भरें।",
    "{label} must be a non-empty string.": "{label} भरना आवश्यक है।",
    "{label} must be one line, at most {limit} characters, with valid text.": "{label} में एक पंक्ति का वैध पाठ दें, अधिकतम {limit} अक्षर।",
    "{label}: use a non-negative decimal (up to 15 integer and 8 fractional digits).": "{label}: शून्य या धनात्मक दशमलव दें (बिंदु से पहले अधिकतम 15 और बाद में 8 अंक)।",
    "Age must be a whole number of years between 1 and 150.": "आयु 1 से 150 के बीच पूर्ण वर्षों में दें।",
    "Document date must be an ISO date string.": "दस्तावेज़ दिनांक YYYY-MM-DD पाठ में दें।",
    "Document date must be a real date in YYYY-MM-DD format.": "दस्तावेज़ की वैध दिनांक YYYY-MM-DD प्रारूप में दें।",
    "Choose legacy_devlys or unicode text mode.": "DevLys 010 या Mangal इनपुट फ़ॉन्ट चुनें।",
    "The Unicode font must be Mangal. DevLys 010 requires pre-encoded text.": "यूनिकोड फ़ॉन्ट Mangal होना चाहिए। DevLys 010 के लिए पहले से एन्कोड किया पाठ आवश्यक है।",
    "Add between 1 and 500 land records before saving or generating.": "सहेजने या दस्तावेज़ बनाने से पहले 1 से 500 भूमि अभिलेख जोड़ें।",
    "Invalid land record.": "अमान्य भूमि अभिलेख।",
    "DevLys mode requires already encoded keyboard text, not Unicode Hindi. No automatic conversion is performed.": "DevLys के लिए पहले से एन्कोड किए कीबोर्ड कोड दें, यूनिकोड हिन्दी नहीं। स्वतः रूपांतरण नहीं किया जाता।",
    "Unsupported or missing JSON schema_version (expected 1 or 2).": "JSON schema_version अनुपस्थित या असमर्थित है (1 या 2 आवश्यक)।",
    "Unknown JSON fields: {fields}": "अज्ञात JSON फ़ील्ड: {fields}",
    "Version 1 JSON cannot contain version 2 client fields.": "संस्करण 1 JSON में संस्करण 2 के ग्राहक फ़ील्ड नहीं हो सकते।",
    "Old JSON has different Bahi numbers. A document requires one Bahi; resolve the numbers before importing.": "पुराने JSON में अलग-अलग बही क्रमांक हैं। एक दस्तावेज़ के लिए एक बही आवश्यक है; आयात से पहले क्रमांक सही करें।",
    "JSON has missing fields or invalid record structure.": "JSON में फ़ील्ड गायब हैं या अभिलेख संरचना अमान्य है।",
    "Client data must be saved to a .json file.": "ग्राहक डेटा .json फ़ाइल में सहेजें।",
    "Output must have a .docx extension.": "दस्तावेज़ का एक्सटेंशन .docx होना चाहिए।",
    "Save generated documents outside templates/.": "बनाए गए दस्तावेज़ templates/ के बाहर सहेजें।",
    "Output already exists. Choose a new filename.": "फ़ाइल पहले से मौजूद है। नया फ़ाइल नाम चुनें।",
    "Template differs from the inspected original. Inspect it and update the engine profile before generating.": "टेम्पलेट जाँचे गए मूल से अलग है। दस्तावेज़ बनाने से पहले टेम्पलेट जाँचें और इंजन प्रोफ़ाइल अद्यतन करें।",
    "Expected all three inspected land tables, including the nested table.": "भीतरी तालिका सहित जाँची गई तीनों भूमि तालिकाएँ आवश्यक हैं।",
    "Technical details: {error}": "तकनीकी विवरण: {error}",
}


def translate(message, language="English", **values):
    result = HINDI.get(message, message) if language == "हिन्दी" else message
    return result.format(**values) if values else result


def error_text(error, language):
    if isinstance(error, ValidationError):
        values = dict(error.values)
        if "label" in values:
            values["label"] = translate(values["label"], language)
        return translate(error.message, language, **values)
    message = str(error)
    if language == "हिन्दी" and message not in HINDI:
        return translate("Technical details: {error}", language, error=message)
    return translate(message, language)

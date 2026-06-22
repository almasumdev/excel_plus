part of '../../excel_plus.dart';

/// Styling class for cells
// ignore: must_be_immutable
class _FontStyle {
  ExcelColor? _fontColorHex = ExcelColor.black;
  String? _fontFamily;
  FontScheme _fontScheme = FontScheme.Unset;
  bool _bold = false, _italic = false;
  Underline _underline = Underline.None;
  int? _fontSize;

  _FontStyle({
    ExcelColor? fontColorHex = ExcelColor.black,
    int? fontSize,
    String? fontFamily,
    FontScheme fontScheme = FontScheme.Unset,
    bool bold = false,
    Underline underline = Underline.None,
    bool italic = false,
  }) {
    _bold = bold;

    _fontSize = fontSize;

    _italic = italic;

    _fontFamily = fontFamily;

    _fontScheme = fontScheme;

    _underline = underline;

    if (fontColorHex != null) {
      _fontColorHex = _appropriateColor(fontColorHex);
    } else {
      _fontColorHex = ExcelColor.black;
    }
  }

  ExcelColor get fontColor {
    return _fontColorHex ?? ExcelColor.black;
  }

  set fontColor(ExcelColor? fontColorHex) {
    if (fontColorHex != null) {
      _fontColorHex = _appropriateColor(fontColorHex);
    } else {
      _fontColorHex = ExcelColor.black;
    }
  }

  String? get fontFamily {
    return _fontFamily;
  }

  set fontFamily(String? family) {
    _fontFamily = family;
  }

  FontScheme get fontScheme {
    return _fontScheme;
  }

  set fontScheme(FontScheme scheme) {
    _fontScheme = scheme;
  }

  int? get fontSize {
    return _fontSize;
  }

  set fontSize(int? fs) {
    _fontSize = fs;
  }

  Underline get underline {
    return _underline;
  }

  set underline(Underline underline) {
    _underline = underline;
  }

  bool get isBold {
    return _bold;
  }

  set isBold(bool bold) {
    _bold = bold;
  }

  bool get isItalic {
    return _italic;
  }

  set isItalic(bool italic) {
    _italic = italic;
  }

  @override
  bool operator ==(Object other) =>
      identical(this, other) ||
      other is _FontStyle &&
          other._bold == _bold &&
          other._italic == _italic &&
          other._fontSize == _fontSize &&
          other._underline == _underline &&
          other._fontFamily == _fontFamily &&
          other._fontColorHex == _fontColorHex &&
          other._fontScheme == _fontScheme;

  @override
  int get hashCode => Object.hash(
    _bold,
    _italic,
    _fontSize,
    _underline,
    _fontFamily,
    _fontColorHex,
    _fontScheme,
  );
}

/// Available font families.
///
/// {@category Styling}
enum FontFamily {
  /// The "Al Bayan Plain" font family.
  Al_Bayan_Plain,

  /// The "Abadi MT Condensed Light" font family.
  Abadi_MT_Condensed_Light,

  /// The "Abadi MT Condensed Extra Bold" font family.
  Abadi_MT_Condensed_Extra_Bold,

  /// The "Al Nile" font family.
  Al_Nile,

  /// The "Al Tarikh Regular" font family.
  Al_Tarikh_Regular,

  /// The "American Typewriter" font family.
  American_Typewriter,

  /// The "Andale Mono" font family.
  Andale_Mono,

  /// The "Angsana New" font family.
  Angsana_New,

  /// The "Apple Braille Outline 8 Dot" font family.
  Apple_Braille_Outline_8_Dot,

  /// The "Apple Chancery" font family.
  Apple_Chancery,

  /// The "Apple Color Emoji" font family.
  Apple_Color_Emoji,

  /// The "Apple Symbols" font family.
  Apple_Symbols,

  /// The "Arial" font family.
  Arial,

  /// The "Arial Hebrew" font family.
  Arial_Hebrew,

  /// The "Arial Hebrew Scholar" font family.
  Arial_Hebrew_Scholar,

  /// The "Arial Narrow" font family.
  Arial_Narrow,

  /// The "Arial Rounded MT Bold" font family.
  Arial_Rounded_MT_Bold,

  /// The "Arial Unicode MS" font family.
  Arial_Unicode_MS,

  /// The "Athelas Regular" font family.
  Athelas_Regular,

  /// The "Avenir Book" font family.
  Avenir_Book,

  /// The "Avenir Next Regular" font family.
  Avenir_Next_Regular,

  /// The "Avenir Next Condensed Regular" font family.
  Avenir_Next_Condensed_Regular,

  /// The "Ayuthaya" font family.
  Ayuthaya,

  /// The "Baghdad" font family.
  Baghdad,

  /// The "Bangla MN" font family.
  Bangla_MN,

  /// The "Bangla Sangam MN" font family.
  Bangla_Sangam_MN,

  /// The "Baskerville" font family.
  Baskerville,

  /// The "Baskerville Old Face" font family.
  Baskerville_Old_Face,

  /// The "Bauhaus 93" font family.
  Bauhaus_93,

  /// The "Beirut" font family.
  Beirut,

  /// The "Bell MT" font family.
  Bell_MT,

  /// The "Bernard MT Condensed" font family.
  Bernard_MT_Condensed,

  /// The "Big Caslon" font family.
  Big_Caslon,

  /// The "Bodoni 72" font family.
  Bodoni_72,

  /// The "Bodoni 72 Oldstyle" font family.
  Bodoni_72_Oldstyle,

  /// The "Bodoni 72 Smallcaps" font family.
  Bodoni_72_Smallcaps,

  /// The "Bodoni Ornaments" font family.
  Bodoni_Ornaments,

  /// The "Book Antiqua" font family.
  Book_Antiqua,

  /// The "Bookman Old Style" font family.
  Bookman_Old_Style,

  /// The "Bookshelf Symbol 7" font family.
  Bookshelf_Symbol_7,

  /// The "Bradley Hand" font family.
  Bradley_Hand,

  /// The "Braggadocio" font family.
  Braggadocio,

  /// The "Britannic Bold" font family.
  Britannic_Bold,

  /// The "Brush Script MT" font family.
  Brush_Script_MT,

  /// The "Calibri" font family.
  Calibri,

  /// The "Calisto MT" font family.
  Calisto_MT,

  /// The "Cambria" font family.
  Cambria,

  /// The "Cambria Math" font family.
  Cambria_Math,

  /// The "Candara" font family.
  Candara,

  /// The "Century" font family.
  Century,

  /// The "Century Gothic" font family.
  Century_Gothic,

  /// The "Century Schoolbook" font family.
  Century_Schoolbook,

  /// The "Chalkboard" font family.
  Chalkboard,

  /// The "Chalkboard SE" font family.
  Chalkboard_SE,

  /// The "Chalkduster" font family.
  Chalkduster,

  /// The "Charter" font family.
  Charter,

  /// The "Cochin" font family.
  Cochin,

  /// The "Colonna MT" font family.
  Colonna_MT,

  /// The "Comic Sans MS" font family.
  Comic_Sans_MS,

  /// The "Consolas" font family.
  Consolas,

  /// The "Constantia" font family.
  Constantia,

  /// The "Cooper Black" font family.
  Cooper_Black,

  /// The "Copperplate" font family.
  Copperplate,

  /// The "Copperplate Gothic Bold" font family.
  Copperplate_Gothic_Bold,

  /// The "Corbel" font family.
  Corbel,

  /// The "Cordia New" font family.
  Cordia_New,

  /// The "CordiaUPC" font family.
  CordiaUPC,

  /// The "Corsiva Hebrew" font family.
  Corsiva_Hebrew,

  /// The "Courier" font family.
  Courier,

  /// The "Courier New" font family.
  Courier_New,

  /// The "Curlz MT" font family.
  Curlz_MT,

  /// The "Damascus" font family.
  Damascus,

  /// The "David" font family.
  David,

  /// The "DecoType Naskh" font family.
  DecoType_Naskh,

  /// The "Desdemona" font family.
  Desdemona,

  /// The "Devanagari MT" font family.
  Devanagari_MT,

  /// The "Devanagari Sangam MN" font family.
  Devanagari_Sangam_MN,

  /// The "Didot" font family.
  Didot,

  /// The "DIN Alternate" font family.
  DIN_Alternate,

  /// The "DIN Condensed" font family.
  DIN_Condensed,

  /// The "Diwan Kufi" font family.
  Diwan_Kufi,

  /// The "Diwan Thuluth" font family.
  Diwan_Thuluth,

  /// The "Dubai" font family.
  Dubai,

  /// The "Edwardian Script ITC" font family.
  Edwardian_Script_ITC,

  /// The "Engravers MT" font family.
  Engravers_MT,

  /// The "Euphemia UCAS" font family.
  Euphemia_UCAS,

  /// The "Eurostile" font family.
  Eurostile,

  /// The "Farah" font family.
  Farah,

  /// The "Farisi" font family.
  Farisi,

  /// The "Footlight MT Light" font family.
  Footlight_MT_Light,

  /// The "Franklin Gothic Book" font family.
  Franklin_Gothic_Book,

  /// The "Franklin Gothic Demi" font family.
  Franklin_Gothic_Demi,

  /// The "Franklin Gothic Demi Cond" font family.
  Franklin_Gothic_Demi_Cond,

  /// The "Franklin Gothic Heavy" font family.
  Franklin_Gothic_Heavy,

  /// The "Franklin Gothic Medium" font family.
  Franklin_Gothic_Medium,

  /// The "Franklin Gothic Medium Cond" font family.
  Franklin_Gothic_Medium_Cond,

  /// The "Futura" font family.
  Futura,

  /// The "Gabriola" font family.
  Gabriola,

  /// The "Galvji" font family.
  Galvji,

  /// The "Garamond" font family.
  Garamond,

  /// The "Gautami" font family.
  Gautami,

  /// The "Geeza Pro" font family.
  Geeza_Pro,

  /// The "Geneva" font family.
  Geneva,

  /// The "Georgia" font family.
  Georgia,

  /// The "Gill Sans" font family.
  Gill_Sans,

  /// The "Gill Sans MT" font family.
  Gill_Sans_MT,

  /// The "Gill Sans MT Condensed" font family.
  Gill_Sans_MT_Condensed,

  /// The "Gill Sans MT Ext Condensed Bold" font family.
  Gill_Sans_MT_Ext_Condensed_Bold,

  /// The "Gill Sans Ultra Bold" font family.
  Gill_Sans_Ultra_Bold,

  /// The "Gloucester MT Extra Condensed" font family.
  Gloucester_MT_Extra_Condensed,

  /// The "Goudy Old Style" font family.
  Goudy_Old_Style,

  /// The "Gujarati MT" font family.
  Gujarati_MT,

  /// The "Gujarati Sangam MN" font family.
  Gujarati_Sangam_MN,

  /// The "Gurmukhi MN" font family.
  Gurmukhi_MN,

  /// The "Gurmukhi MT" font family.
  Gurmukhi_MT,

  /// The "Gurmukhi Sangam MN" font family.
  Gurmukhi_Sangam_MN,

  /// The "Haettenschweiler" font family.
  Haettenschweiler,

  /// The "Harrington" font family.
  Harrington,

  /// The "Helvetica" font family.
  Helvetica,

  /// The "Helvetica Neue" font family.
  Helvetica_Neue,

  /// The "Herculanum" font family.
  Herculanum,

  /// The "Hoefler Text" font family.
  Hoefler_Text,

  /// The "Impact" font family.
  Impact,

  /// The "Imprint MT Shadow" font family.
  Imprint_MT_Shadow,

  /// The "InaiMathi" font family.
  InaiMathi,

  /// The "Iowan Old Style" font family.
  Iowan_Old_Style,

  /// The "ITF Devanagari" font family.
  ITF_Devanagari,

  /// The "ITF Devanagari Marathi" font family.
  ITF_Devanagari_Marathi,

  /// The "Kailasa" font family.
  Kailasa,

  /// The "Kannada MN" font family.
  Kannada_MN,

  /// The "Kannada Sangam MN" font family.
  Kannada_Sangam_MN,

  /// The "Kartika" font family.
  Kartika,

  /// The "Kefa" font family.
  Kefa,

  /// The "Khmer MN" font family.
  Khmer_MN,

  /// The "Khmer Sangam MN" font family.
  Khmer_Sangam_MN,

  /// The "Kino MT" font family.
  Kino_MT,

  /// The "Kohinoor Bangla" font family.
  Kohinoor_Bangla,

  /// The "Kohinoor Devanagari" font family.
  Kohinoor_Devanagari,

  /// The "Kohinoor Gujarati" font family.
  Kohinoor_Gujarati,

  /// The "Kohinoor Telugu" font family.
  Kohinoor_Telugu,

  /// The "Kokonor" font family.
  Kokonor,

  /// The "Lao MN" font family.
  Lao_MN,

  /// The "Lao Sangam MN" font family.
  Lao_Sangam_MN,

  /// The "Latha" font family.
  Latha,

  /// The "Lucida Blackletter" font family.
  Lucida_Blackletter,

  /// The "Lucida Bright" font family.
  Lucida_Bright,

  /// The "Lucida Calligraphy" font family.
  Lucida_Calligraphy,

  /// The "Lucida Console" font family.
  Lucida_Console,

  /// The "Lucida Fax" font family.
  Lucida_Fax,

  /// The "Lucida Grande" font family.
  Lucida_Grande,

  /// The "Lucida Handwriting" font family.
  Lucida_Handwriting,

  /// The "Lucida Sans" font family.
  Lucida_Sans,

  /// The "Lucida Sans Typewriter" font family.
  Lucida_Sans_Typewriter,

  /// The "Lucida Sans Unicode" font family.
  Lucida_Sans_Unicode,

  /// The "Luminari" font family.
  Luminari,

  /// The "Malayalam MN" font family.
  Malayalam_MN,

  /// The "Malayalam Sangam MN" font family.
  Malayalam_Sangam_MN,

  /// The "Mangal" font family.
  Mangal,

  /// The "Marion" font family.
  Marion,

  /// The "Marker Felt" font family.
  Marker_Felt,

  /// The "Marlett" font family.
  Marlett,

  /// The "Matura MT Script Capitals" font family.
  Matura_MT_Script_Capitals,

  /// The "Menlo" font family.
  Menlo,

  /// The "Microsoft New Tai Lue" font family.
  Microsoft_New_Tai_Lue,

  /// The "Microsoft Sans Serif" font family.
  Microsoft_Sans_Serif,

  /// The "Microsoft Tai Le" font family.
  Microsoft_Tai_Le,

  /// The "Microsoft Yi Baiti" font family.
  Microsoft_Yi_Baiti,

  /// The "Mishafi" font family.
  Mishafi,

  /// The "Mishafi Gold" font family.
  Mishafi_Gold,

  /// The "Mistral" font family.
  Mistral,

  /// The "Monaco" font family.
  Monaco,

  /// The "Monotype Corsiva" font family.
  Monotype_Corsiva,

  /// The "Monotype Sorts" font family.
  Monotype_Sorts,

  /// The "MS Reference Sans Serif" font family.
  MS_Reference_Sans_Serif,

  /// The "MS Reference Specialty" font family.
  MS_Reference_Specialty,

  /// The "Mshtakan" font family.
  Mshtakan,

  /// The "MT Extra" font family.
  MT_Extra,

  /// The "Mukta Mahee" font family.
  Mukta_Mahee,

  /// The "Muna" font family.
  Muna,

  /// The "Myanmar MN" font family.
  Myanmar_MN,

  /// The "Myanmar Sangam MN" font family.
  Myanmar_Sangam_MN,

  /// The "Myanmar Text" font family.
  Myanmar_Text,
}

///
///
///returns the `Font Family Name`
///
///
/// @nodoc
String getFontFamily(FontFamily fontFamily) {
  return (fontFamily.toString().replaceAll(
    'FontFamily.',
    '',
  )).replaceAll('_', ' ');
}

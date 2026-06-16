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
  Al_Bayan_Plain,
  Abadi_MT_Condensed_Light,
  Abadi_MT_Condensed_Extra_Bold,
  Al_Nile,
  Al_Tarikh_Regular,
  American_Typewriter,
  Andale_Mono,
  Angsana_New,
  Apple_Braille_Outline_8_Dot,
  Apple_Chancery,
  Apple_Color_Emoji,
  Apple_Symbols,
  Arial,
  Arial_Hebrew,
  Arial_Hebrew_Scholar,
  Arial_Narrow,
  Arial_Rounded_MT_Bold,
  Arial_Unicode_MS,
  Athelas_Regular,
  Avenir_Book,
  Avenir_Next_Regular,
  Avenir_Next_Condensed_Regular,
  Ayuthaya,
  Baghdad,
  Bangla_MN,
  Bangla_Sangam_MN,
  Baskerville,
  Baskerville_Old_Face,
  Bauhaus_93,
  Beirut,
  Bell_MT,
  Bernard_MT_Condensed,
  Big_Caslon,
  Bodoni_72,
  Bodoni_72_Oldstyle,
  Bodoni_72_Smallcaps,
  Bodoni_Ornaments,
  Book_Antiqua,
  Bookman_Old_Style,
  Bookshelf_Symbol_7,
  Bradley_Hand,
  Braggadocio,
  Britannic_Bold,
  Brush_Script_MT,
  Calibri,
  Calisto_MT,
  Cambria,
  Cambria_Math,
  Candara,
  Century,
  Century_Gothic,
  Century_Schoolbook,
  Chalkboard,
  Chalkboard_SE,
  Chalkduster,
  Charter,
  Cochin,
  Colonna_MT,
  Comic_Sans_MS,
  Consolas,
  Constantia,
  Cooper_Black,
  Copperplate,
  Copperplate_Gothic_Bold,
  Corbel,
  Cordia_New,
  CordiaUPC,
  Corsiva_Hebrew,
  Courier,
  Courier_New,
  Curlz_MT,
  Damascus,
  David,
  DecoType_Naskh,
  Desdemona,
  Devanagari_MT,
  Devanagari_Sangam_MN,
  Didot,
  DIN_Alternate,
  DIN_Condensed,
  Diwan_Kufi,
  Diwan_Thuluth,
  Dubai,
  Edwardian_Script_ITC,
  Engravers_MT,
  Euphemia_UCAS,
  Eurostile,
  Farah,
  Farisi,
  Footlight_MT_Light,
  Franklin_Gothic_Book,
  Franklin_Gothic_Demi,
  Franklin_Gothic_Demi_Cond,
  Franklin_Gothic_Heavy,
  Franklin_Gothic_Medium,
  Franklin_Gothic_Medium_Cond,
  Futura,
  Gabriola,
  Galvji,
  Garamond,
  Gautami,
  Geeza_Pro,
  Geneva,
  Georgia,
  Gill_Sans,
  Gill_Sans_MT,
  Gill_Sans_MT_Condensed,
  Gill_Sans_MT_Ext_Condensed_Bold,
  Gill_Sans_Ultra_Bold,
  Gloucester_MT_Extra_Condensed,
  Goudy_Old_Style,
  Gujarati_MT,
  Gujarati_Sangam_MN,
  Gurmukhi_MN,
  Gurmukhi_MT,
  Gurmukhi_Sangam_MN,
  Haettenschweiler,
  Harrington,
  Helvetica,
  Helvetica_Neue,
  Herculanum,
  Hoefler_Text,
  Impact,
  Imprint_MT_Shadow,
  InaiMathi,
  Iowan_Old_Style,
  ITF_Devanagari,
  ITF_Devanagari_Marathi,
  Kailasa,
  Kannada_MN,
  Kannada_Sangam_MN,
  Kartika,
  Kefa,
  Khmer_MN,
  Khmer_Sangam_MN,
  Kino_MT,
  Kohinoor_Bangla,
  Kohinoor_Devanagari,
  Kohinoor_Gujarati,
  Kohinoor_Telugu,
  Kokonor,
  Lao_MN,
  Lao_Sangam_MN,
  Latha,
  Lucida_Blackletter,
  Lucida_Bright,
  Lucida_Calligraphy,
  Lucida_Console,
  Lucida_Fax,
  Lucida_Grande,
  Lucida_Handwriting,
  Lucida_Sans,
  Lucida_Sans_Typewriter,
  Lucida_Sans_Unicode,
  Luminari,
  Malayalam_MN,
  Malayalam_Sangam_MN,
  Mangal,
  Marion,
  Marker_Felt,
  Marlett,
  Matura_MT_Script_Capitals,
  Menlo,
  Microsoft_New_Tai_Lue,
  Microsoft_Sans_Serif,
  Microsoft_Tai_Le,
  Microsoft_Yi_Baiti,
  Mishafi,
  Mishafi_Gold,
  Mistral,
  Monaco,
  Monotype_Corsiva,
  Monotype_Sorts,
  MS_Reference_Sans_Serif,
  MS_Reference_Specialty,
  Mshtakan,
  MT_Extra,
  Mukta_Mahee,
  Muna,
  Myanmar_MN,
  Myanmar_Sangam_MN,
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

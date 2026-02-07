use rust_xlsxwriter::{Color, Format, FormatBorder};

pub(crate) enum DataType {
    Money,     // Rounded to 2 decimal places and formatted as currency
    UsualText, // Nothing to do
}

pub(crate) enum CellType {
    Usual,       // For usual cells - white background dotted border and bold font
    Weekend,     // For weekend day cells - white red background dotted border and bold font
    Earn,        // For earn cells - green background dotted border and bold font
    Header,      // For header cells - pink background and solid border normal font
    TotalBonus,  // For total bonus cells - white background solid border and bold font
    InputHeader, // For input header cells - orange background solid border and normal font
    MonthWinter, // For month winter cells - blue background solid border and normal font
    MonthSummer, // For month summer cells - yellow background solid border and normal font
    MonthAutumn, // For month autumn cells - orange background solid border and normal font
    MonthSpring, // For month spring cells - green background solid border and normal font
}

pub(crate) fn cell_style(data_type: DataType, cell_type: CellType) -> Format {
    let format = match cell_type {
        CellType::Usual => Format::new().set_border(FormatBorder::Dotted).set_bold(),

        CellType::Weekend => Format::new()
            .set_border(FormatBorder::Dotted)
            .set_bold()
            .set_background_color(Color::RGB(0xF8B9B8)),

        CellType::Earn => Format::new()
            .set_border(FormatBorder::Dotted)
            .set_bold()
            .set_background_color(Color::RGB(0xB2E19E)),

        CellType::Header => Format::new()
            .set_border(FormatBorder::Medium)
            .set_background_color(Color::RGB(0xEDCAE9)),

        CellType::TotalBonus => Format::new().set_border(FormatBorder::Medium).set_bold(),

        CellType::InputHeader => Format::new()
            .set_border(FormatBorder::Medium)
            .set_background_color(Color::RGB(0xF0C2A7)),

        CellType::MonthWinter => Format::new()
            .set_border(FormatBorder::Medium)
            .set_background_color(Color::RGB(0xC6E8F4)),

        CellType::MonthSpring => Format::new()
            .set_border(FormatBorder::Medium)
            .set_background_color(Color::RGB(0xB2E19E)),

        CellType::MonthSummer => Format::new()
            .set_border(FormatBorder::Medium)
            .set_background_color(Color::RGB(0xFABD02)),

        CellType::MonthAutumn => Format::new()
            .set_border(FormatBorder::Medium)
            .set_background_color(Color::RGB(0xF0C1A7)),
    };
    format
}

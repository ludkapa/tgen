use anyhow::Result as AResult;
use rust_xlsxwriter::Format;

pub(crate) enum DataType {
    Money,     // Rounded to 2 decimal places and formatted as currency
    UsualText, // Nothing to do
}

pub(crate) enum CellType {
    Usual,       // For usual cells - white background dotted border and bold font
    Weekend,     // For weekend day cells - white red background dotted border and bold font
    Earn,        // For earn cells - green background dotted border and bold font
    Headers,     // For header cells - pink background and solid border normal font
    TotalBonus,  // For total bonus cells - white background solid border and bold font
    InputHeader, // For input header cells - orange background solid border and normal font
    MonthWinter, // For month winter cells - blue background solid border and normal font
    MonthSummer, // For month summer cells - yellow background solid border and normal font
    MonthAutumn, // For month autumn cells - orange background solid border and normal font
    MonthSpring, // For month spring cells - green background solid border and normal font
}

pub(crate) fn CellStyle(data_type: DataType, cell_type: CellType) -> AResult<Format> {
    let format = Format::new();
    todo!()
}

requires 'Moose';
requires 'Excel::Writer::XLSX';
requires 'Statistics::Descriptive';
requires 'Chart::Math::Axis';
requires 'YAML::Syck';
requires 'List::MoreUtils';
requires 'perl', '5.008001';

on test => sub {
    requires 'Test::More', 0.88;
    requires 'Spreadsheet::XLSX', '0.15';
};

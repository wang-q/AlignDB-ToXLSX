requires 'Moose';
requires 'Excel::Writer::XLSX', '0.95';
requires 'Statistics::Descriptive';
requires 'Chart::Math::Axis';
requires 'YAML::Syck', '1.29';
requires 'List::MoreUtils';
requires 'perl', '5.010001';

on test => sub {
    requires 'Test::More', 0.88;
    requires 'Spreadsheet::XLSX', '0.15';
};

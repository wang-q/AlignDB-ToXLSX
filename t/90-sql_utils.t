use strict;
use warnings;
use Test::More;

use Path::Tiny;
use Spreadsheet::XLSX;
use DBI;

use AlignDB::ToXLSX;

# cd ~/Scripts/alignDB
# perl util/query_sql.pl -d ScervsRM11_1a_Spar -t csv -o isw.csv \
#     -q "SELECT isw_id, isw_distance, isw_pi FROM isw LIMIT 1000"

my $temp = Path::Tiny->tempfile;

{
    #@type DBI
    my $dbh = DBI->connect("DBI:CSV:");
    $dbh->{csv_tables}->{isw} = {
        eol            => "\n",
        sep_char       => ",",
        file           => "t/isw.csv",
        skip_first_row => 1,
        quote_char     => '',
        col_names      => [qw{ isw_id isw_length isw_distance isw_pi }],
    };

    my $toxlsx = AlignDB::ToXLSX->new(
        dbh     => $dbh,
        outfile => $temp->stringify,

        #        outfile => "03.xlsx",
    );

    {    # make_combine
        my $sql_query = q{
            SELECT
                isw.isw_distance distance,
                COUNT(*) COUNT,
                SUM(isw.isw_length) SUM_length
            FROM
                isw
            WHERE
                isw.isw_distance <= 100
            GROUP BY
                distance
            ORDER BY
                distance
        };

        my $combined = $toxlsx->make_combine(
            {   sql_query  => $sql_query,
                threshold  => 100,
                standalone => [ -1, 0 ],
                merge_last => 1,
            }
        );
        is( scalar( @{$combined} ), 10, "combined pieces" );
    }
}

#{
#    my $xlsx  = Spreadsheet::XLSX->new( $temp->stringify );
#    my $sheet = $xlsx->{Worksheet}[0];
#
#    is( $sheet->{Name},              "d1_pi", "Sheet Name" );
#    is( $sheet->{MaxRow},            22,      "Sheet MaxRow" );
#    is( $sheet->{MaxCol},            2,       "Sheet MaxCol" );
#    is( $sheet->{Cells}[0][2]{Val},  "COUNT", "Cell content 1" );
#    is( $sheet->{Cells}[1][0]{Val},  -1,      "Cell content 2" );
#    is( $sheet->{Cells}[22][2]{Val}, 15,      "Cell content 3" );
#}

ok(1);

done_testing();

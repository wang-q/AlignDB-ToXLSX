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
        col_names      => [qw{ isw_id isw_distance isw_pi }],
    };

    my $sql_query = q{
        SELECT
            isw.isw_distance distance,
            AVG(isw.isw_pi) AVG_pi,
            COUNT(*) COUNT
        FROM
            isw
        WHERE
            isw.isw_distance <= 20
        GROUP BY
            distance
        ORDER BY
            distance
    };

    my $toxlsx = AlignDB::ToXLSX->new(
        dbh     => $dbh,
        outfile => $temp->stringify,

        #        outfile => "03.xlsx",
    );

    my $sheet_name = 'd1_pi';
    my $sheet;

    {    # header
        my @names = $toxlsx->sql2names($sql_query);
        $sheet = $toxlsx->write_header( $sheet_name, { header => \@names, } );
    }

    my $data;
    {    # content
        $data = $toxlsx->write_sql(
            $sheet,
            {   sql_query => $sql_query,
                data      => 1,
            }
        );
    }

}

{
    my $xlsx  = Spreadsheet::XLSX->new( $temp->stringify );
    my $sheet = $xlsx->{Worksheet}[0];

    is( $sheet->{Name},              "d1_pi", "Sheet Name" );
    is( $sheet->{MaxRow},            22,      "Sheet MaxRow" );
    is( $sheet->{MaxCol},            2,       "Sheet MaxCol" );
    is( $sheet->{Cells}[0][2]{Val},  "COUNT", "Cell content 1" );
    is( $sheet->{Cells}[1][0]{Val},  -1,      "Cell content 2" );
    is( $sheet->{Cells}[22][2]{Val}, 15,      "Cell content 3" );
}

ok(1);

done_testing();

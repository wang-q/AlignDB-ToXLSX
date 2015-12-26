package AlignDB::ToXLSX;
use strict;
use warnings;
use autodie;

use 5.008001;

our $VERSION = '1.0.1';

# ABSTRACT: Generate xlsx files from SQL queries or just arrays.

use Moose;
use Carp;

use Excel::Writer::XLSX;
use DBI;
use Statistics::Descriptive;
use Chart::Math::Axis;
use List::Util qw(first max maxstr min minstr reduce shuffle sum);
use List::MoreUtils qw( all any );

use YAML qw(Dump Load DumpFile LoadFile);

# mysql
has 'mysql'  => ( is => 'ro', isa => 'Str' );    # e.g. 'alignDB:202.119.43.5'
has 'server' => ( is => 'ro', isa => 'Str' );    # e.g. '202.119.43.5'
has 'db'     => ( is => 'ro', isa => 'Str' );    # e.g. 'alignDB'
has 'user'   => ( is => 'ro', isa => 'Str' );    # database username
has 'passwd' => ( is => 'ro', isa => 'Str' );    # database password
has 'dbh'    => ( is => 'ro', isa => 'Ref' );    # store database handle here

has 'mocking' => ( is => 'ro', isa => 'Bool', default => sub {0}, );    # don't connect to mysql

# outfiles
has 'outfile'  => ( is => 'ro', isa => 'Str' );                         # output file, autogenerable
has 'workbook' => ( is => 'rw', isa => 'Object' );                      # excel workbook object
has 'format'   => ( is => 'ro', isa => 'HashRef' );                     # excel formats

# charts
has 'font_name' => ( is => 'rw', isa => 'Str', default => sub {'Arial'}, );
has 'font_size' => ( is => 'rw', isa => 'Num', default => sub {10}, );
has 'width'     => ( is => 'rw', isa => 'Num', default => sub {320}, );
has 'height'    => ( is => 'rw', isa => 'Num', default => sub {320}, );
has 'max_ticks' => ( is => 'rw', isa => 'Int', default => sub {6} );

# Replace texts in titles
has 'replace' => ( is => 'rw', isa => 'HashRef', default => sub { {} } );

sub BUILD {
    my $self = shift;

    # Connect to database
    if ( $self->mysql ) {
        my ( $server, $db ) = split ':', $self->mysql;
        $self->{server} ||= $server;
        $self->{db}     ||= $db;
    }
    elsif ( $self->server and $self->db ) {
        $self->{mysql} = $self->db . ':' . $self->server;
    }
    elsif ( $self->mocking ) {

        # do nothing
    }
    else {
        confess "You should provide either mysql or db-server\n";
    }

    my $mysql  = $self->mysql;
    my $user   = $self->user;
    my $passwd = $self->passwd;
    my $server = $self->server;
    my $db     = $self->db;

    my $dbh = {};
    if ( !$self->mocking ) {
        $dbh = DBI->connect( "dbi:mysql:$mysql", $user, $passwd )
            or confess "Cannot connect to MySQL database at $mysql";
    }
    $self->{dbh} = $dbh;

    # set outfile
    unless ( $self->outfile ) {
        $self->mysql =~ /^(.+):/;
        $self->{outfile} = "$1.auto.xlsx";
    }

    # Create $workbook object
    my $workbook;
    unless ( $workbook = Excel::Writer::XLSX->new( $self->outfile ) ) {
        warn "Cannot create Excel file.\n";
        return;
    }
    $self->{workbook} = $workbook;

    # set $workbook format
    my %font = (
        font => $self->font_name,
        size => $self->font_size,
    );
    my $format = {
        HEADER => $workbook->add_format(
            align    => 'center',
            bg_color => 42,
            bold     => 1,
            bottom   => 2,
            %font,
        ),
        HIGHLIGHT => $workbook->add_format( color => 'blue',  %font, ),
        NORMAL    => $workbook->add_format( color => 'black', %font, ),
        NAME      => $workbook->add_format( bold  => 1,       color => 57, %font, ),
        TOTAL     => $workbook->add_format( bold  => 1,       top => 2, %font, ),
        DATE => $workbook->add_format(
            bg_color   => 42,
            bold       => 1,
            num_format => 'yy-m-d hh:mm',
            %font,
        ),
    };
    $self->{format} = $format;

    return;
}

sub write_header_direct {
    my ( $self, $sheet_name, $option ) = @_;

    # init
    my $workbook = $self->workbook;
    my $sheet    = $workbook->add_worksheet($sheet_name);
    my $fmt      = $self->format;

    # init table cursor
    my $sheet_row = $option->{sheet_row};
    my $sheet_col = $option->{sheet_col};
    my $header    = $option->{header};

    # query name
    my $query_name = $option->{query_name};

    # create table header
    for ( my $i = 0; $i < $sheet_col; $i++ ) {
        $sheet->write( $sheet_row, $i, $query_name, $fmt->{HEADER} );
    }
    for ( my $i = 0; $i < scalar @{$header}; $i++ ) {
        $sheet->write( $sheet_row, $i + $sheet_col, $header->[$i], $fmt->{HEADER} );
    }
    $sheet_row++;
    $sheet->freeze_panes( 1, 0 );    # freeze table

    return ( $sheet, $sheet_row );
}

sub write_header_sql {
    my ( $self, $sheet_name, $option ) = @_;

    # init
    my $dbh      = $self->dbh;
    my $workbook = $self->workbook;
    my $sheet    = $workbook->add_worksheet($sheet_name);
    my $fmt      = $self->format;

    # init table cursor
    my $sheet_row = $option->{sheet_row};
    my $sheet_col = $option->{sheet_col};

    # query name
    my $query_name = $option->{query_name};

    # init DBI query
    my $sql_query = $option->{sql_query};
    my $sth       = $dbh->prepare($sql_query);
    $sth->execute();

    # create table header
    my @cols_name = @{ $sth->{'NAME'} };
    for ( my $i = 0; $i < $sheet_col; $i++ ) {
        $sheet->write( $sheet_row, $i, $query_name, $fmt->{HEADER} );
    }
    for ( my $i = 0; $i < scalar @cols_name; $i++ ) {
        $sheet->write( $sheet_row, $i + $sheet_col, $cols_name[$i], $fmt->{HEADER} );
    }
    $sheet_row++;
    $sheet->freeze_panes( 1, 0 );    # freeze table

    return ( $sheet, $sheet_row );
}

sub write_row_direct {
    my ( $self, $sheet, $option ) = @_;

    # init
    my $fmt = $self->format;

    # init table cursor
    my $sheet_row = $option->{sheet_row};
    my $sheet_col = $option->{sheet_col};

    # query name
    my $query_name = $option->{query_name};
    if ( defined $query_name ) {
        $sheet->write( $sheet_row, $sheet_col - 1, $query_name, $fmt->{NAME} );
    }

    # array_ref
    my $row = $option->{row};

    # content format
    my $content_format = $option->{content_format};
    unless ( defined $content_format ) {
        $content_format = 'NORMAL';
    }

    # reverse write
    my $write_step = $option->{write_step};
    unless ( defined $write_step ) {
        $write_step = 1;
    }

    # bind value
    my $append_column = $option->{append_column};
    unless ( defined $append_column ) {
        $append_column = [];
    }

    # insert table columns
    for ( my $i = 0; $i < scalar @$row; $i++ ) {
        $sheet->write( $sheet_row, $i + $sheet_col, $row->[$i], $fmt->{$content_format} );
    }
    $sheet_row += $write_step;

    return ($sheet_row);
}

sub write_content_direct {
    my ( $self, $sheet, $option ) = @_;

    # init
    my $dbh = $self->dbh;
    my $fmt = $self->format;

    # init table cursor
    my $sheet_row = $option->{sheet_row};
    my $sheet_col = $option->{sheet_col};

    # query name
    my $query_name = $option->{query_name};
    if ( defined $query_name ) {
        $sheet->write( $sheet_row, $sheet_col - 1, $query_name, $fmt->{NAME} );
    }

    # content format
    my $content_format = $option->{content_format};
    unless ( defined $content_format ) {
        $content_format = 'NORMAL';
    }

    # bind value
    my $bind_value = $option->{bind_value};
    unless ( defined $bind_value ) {
        $bind_value = [];
    }

    # reverse write
    my $write_step = $option->{write_step};
    unless ( defined $write_step ) {
        $write_step = 1;
    }

    # append column
    my $append_column = $option->{append_column};
    unless ( defined $append_column ) {
        $append_column = [];
    }

    # init DBI query
    my $sql_query = $option->{sql_query};
    my $sth       = $dbh->prepare($sql_query);
    $sth->execute(@$bind_value);

    # insert table columns
    while ( my @row = $sth->fetchrow_array ) {
        for ( my $i = 0; $i < scalar @row; $i++ ) {
            $sheet->write( $sheet_row, $i + $sheet_col, $row[$i], $fmt->{$content_format} );
        }
        if ( scalar @$append_column ) {
            my $appand_row = shift @$append_column;
            for ( my $i = 0; $i < scalar @$appand_row; $i++ ) {
                $sheet->write(
                    $sheet_row,        $i + $sheet_col + scalar(@row),
                    $appand_row->[$i], $fmt->{$content_format}
                );
            }
        }
        $sheet_row += $write_step;
    }

    return ($sheet_row);
}

sub write_content_combine {
    my ( $self, $sheet, $option ) = @_;

    # init table cursor
    my $sheet_row = $option->{sheet_row};
    my $sheet_col = $option->{sheet_col};
    my $sql_query = $option->{sql_query};

    my @combined = @{ $option->{combined} };

    # bind value
    my $bind_value = $option->{bind_value};
    unless ( defined $bind_value ) {
        $bind_value = [];
    }

    foreach (@combined) {
        my @range      = @$_;
        my $in_list    = '(' . join( ',', @range ) . ')';
        my $sql_query2 = $sql_query . $in_list;
        my %option     = (
            sql_query  => $sql_query2,
            sheet_row  => $sheet_row,
            sheet_col  => $sheet_col,
            bind_value => $bind_value,
        );
        ($sheet_row) = $self->write_content_direct( $sheet, \%option );
    }
    return ($sheet_row);
}

sub write_content_group {
    my ( $self, $sheet, $option ) = @_;

    # init table cursor
    my $sheet_row     = $option->{sheet_row};
    my $sheet_col     = $option->{sheet_col};
    my $sql_query     = $option->{sql_query};
    my $append_column = $option->{append_column};

    # bind value
    my $bind_value = $option->{bind_value};
    unless ( defined $bind_value ) {
        $bind_value = [];
    }

    my @group = @{ $option->{group} };

    foreach (@group) {
        my @range      = @$_;
        my $in_list    = '(' . join( ',', @range ) . ')';
        my $sql_query2 = $sql_query . $in_list;
        my $group_name;
        if ( scalar @range > 1 ) {
            $group_name = $range[0] . "--" . $range[-1];
        }
        else {
            $group_name = $range[0];
        }
        my %option = (
            sql_query     => $sql_query2,
            sheet_row     => $sheet_row,
            sheet_col     => $sheet_col,
            query_name    => $group_name,
            append_column => $append_column,
            bind_value    => $bind_value,
        );
        ($sheet_row) = $self->write_content_direct( $sheet, \%option );
    }
    return ($sheet_row);
}

sub write_content_series {
    my ( $self, $sheet, $option ) = @_;

    # init objects
    my $dbh = $self->dbh;
    my $fmt = $self->format;

    # init table cursor
    my $sheet_row = $option->{sheet_row};
    my $sheet_col = $option->{sheet_col};

    my $sql_query = $option->{sql_query};
    my @group     = @{ $option->{group} };

    foreach (@group) {
        my @range = @$_;
        my $group_name;
        if ( scalar @range > 1 ) {
            $group_name = join "-", @range;
        }
        else {
            $group_name = $range[0];
        }
        $sheet_row++;    # add a blank line
        my %option = (
            sql_query  => $sql_query,
            sheet_row  => $sheet_row,
            sheet_col  => $sheet_col,
            query_name => $group_name,
            bind_value => \@range,
        );
        ($sheet_row) = $self->write_content_direct( $sheet, \%option );
    }
}

sub write_content_highlight {
    my ( $self, $sheet, $option ) = @_;

    # init
    my $dbh = $self->dbh;
    my $fmt = $self->format;

    # init table cursor
    my $sheet_row = $option->{sheet_row};
    my $sheet_col = $option->{sheet_col};

    # bind value
    my $bind_value = $option->{bind_value};
    unless ( defined $bind_value ) {
        $bind_value = [];
    }

    # init DBI query
    my $sql_query = $option->{sql_query};
    my $sth       = $dbh->prepare($sql_query);
    $sth->execute(@$bind_value);

    # insert table columns
    my $last_number;
    while ( my @row = $sth->fetchrow_array ) {

        # Highlight 'special' indels
        my $style = 'NORMAL';
        if ( defined $last_number ) {
            if ( $row[1] > $last_number ) {
                $style = 'HIGHLIGHT';
            }
        }
        $last_number = $row[1];
        for ( my $i = 0; $i < scalar @row; $i++ ) {
            $sheet->write( $sheet_row, $i + $sheet_col, $row[$i], $fmt->{$style} );
        }
        $sheet_row++;
    }
}

sub write_content_column {
    my ( $self, $sheet, $option ) = @_;

    # init objects
    my $dbh = $self->dbh;
    my $fmt = $self->format;

    # init table cursor
    my $sheet_row = $option->{sheet_row};
    my $sheet_col = $option->{sheet_col};

    my $sql_query  = $option->{sql_query};
    my @conditions = @{ $option->{conditions} };

    for ( my $i = 0; $i < scalar @conditions; $i++ ) {
        my @bind_values = @{ $conditions[$i] };

        my $sub_sheet_row = $sheet_row;

        my $sth = $dbh->prepare($sql_query);
        $sth->execute(@bind_values);

        # insert table columns
        while ( my @row = $sth->fetchrow_array ) {
            $sheet->write( $sub_sheet_row, 0 + $sheet_col,      $row[0], $fmt->{NORMAL} );
            $sheet->write( $sub_sheet_row, $i + 1 + $sheet_col, $row[1], $fmt->{NORMAL} );
            $sub_sheet_row++;
        }
        $sth->finish;
    }
}

sub make_combine {
    my ( $self, $option ) = @_;

    # init objects
    my $dbh = $self->dbh;

    # init parameters
    my $sql_query  = $option->{sql_query};
    my $threshold  = $option->{threshold};
    my $standalone = $option->{standalone};

    # bind value
    my $bind_value = $option->{bind_value};
    unless ( defined $bind_value ) {
        $bind_value = [];
    }

    # merge_last
    my $merge_last = $option->{merge_last};
    unless ( defined $merge_last ) {
        $merge_last = 0;
    }

    # init DBI query
    my $sth = $dbh->prepare($sql_query);
    $sth->execute(@$bind_value);

    my @row_count = ();
    while ( my @row = $sth->fetchrow_array ) {
        push @row_count, \@row;
    }

    my @combined;    # return these
    my @temp_combined = ();
    my $temp_count    = 0;
    foreach my $row_ref (@row_count) {
        if ( any { $_ eq $row_ref->[0] } @{$standalone} ) {
            push @combined, [ $row_ref->[0] ];
        }
        elsif ( $temp_count < $threshold ) {
            push @temp_combined, $row_ref->[0];
            $temp_count += $row_ref->[1];

            if ( $temp_count < $threshold ) {
                next;
            }
            else {
                push @combined, [@temp_combined];
                @temp_combined = ();
                $temp_count    = 0;
            }
        }
        else {
            warn "Errors occured in calculating combined distance.\n";
        }
    }

    # Write the last weighted row which COUNT might
    #   be smaller than $threshold
    if ( $temp_count > 0 ) {
        if ($merge_last) {
            if ( @combined == 0 ) {
                @combined = ( [] );
            }
            push @{ $combined[-1] }, @temp_combined;
        }
        else {
            push @combined, [@temp_combined];
        }
    }

    return \@combined;
}

sub make_combine_piece {
    my ( $self, $option ) = @_;

    # init objects
    my $dbh = $self->dbh;

    # init parameters
    my $sql_query = $option->{sql_query};
    my $piece     = $option->{piece};

    # bind value
    my $bind_value = $option->{bind_value};
    unless ( defined $bind_value ) {
        $bind_value = [];
    }

    # init DBI query
    my $sth = $dbh->prepare($sql_query);
    $sth->execute(@$bind_value);

    my @row_count = ();
    while ( my @row = $sth->fetchrow_array ) {
        push @row_count, \@row;
    }

    my $sum;
    $sum += $_->[1] for @row_count;
    my $small_chunk = $sum / $piece;

    my @combined;    # return these
    my @temp_combined = ();
    my $temp_count    = 0;
    for my $row_ref (@row_count) {
        if ( $temp_count < $small_chunk ) {
            push @temp_combined, $row_ref->[0];
            $temp_count += $row_ref->[1];

            if ( $temp_count >= $small_chunk ) {
                push @combined, [@temp_combined];
                @temp_combined = ();
                $temp_count    = 0;
            }
        }
        else {
            warn "Errors occured in calculating combined distance.\n";
        }
    }

    # Write the last weighted row which COUNT might
    #   be smaller than $threshold
    if ( $temp_count > 0 ) {
        push @combined, [@temp_combined];
    }

    return \@combined;
}

sub make_last_portion {
    my ( $self, $option ) = @_;

    # init objects
    my $dbh = $self->dbh;

    # init parameters
    my $sql_query = $option->{sql_query};
    my $portion   = $option->{portion};

    # init DBI query
    my $sth = $dbh->prepare($sql_query);
    $sth->execute;

    my @row_count = ();
    while ( my @row = $sth->fetchrow_array ) {
        push @row_count, \@row;
    }

    my @last_portion;    # return @last_portion
    my $all_length = 0;  # return $all_length
    foreach (@row_count) {
        $all_length += $_->[2];
    }
    my @rev_row_count = reverse @row_count;
    my $temp_length   = 0;
    foreach (@rev_row_count) {
        push @last_portion, $_->[0];
        $temp_length += $_->[2];
        if ( $temp_length >= $all_length * $portion ) {
            last;
        }
    }

    return ( $all_length, \@last_portion );
}

sub excute_sql {
    my ( $self, $option ) = @_;

    # init
    my $dbh = $self->dbh;

    # bind value
    my $bind_value = $option->{bind_value};
    unless ( defined $bind_value ) {
        $bind_value = [];
    }

    # init DBI query
    my $sql_query = $option->{sql_query};
    my $sth       = $dbh->prepare($sql_query);
    $sth->execute(@$bind_value);
}

sub check_column {
    my ( $self, $table, $column ) = @_;

    # init
    my $dbh = $self->dbh;

    {    # check table existing
        my @table_names = $dbh->tables( '', '', '' );

        # table names are quoted by ` (back-quotes) which is the
        #   quote_identifier
        my $table_name = "`$table`";
        unless ( any { $_ =~ /$table_name/i } @table_names ) {
            print " " x 4, "Table $table does not exist\n";
            return 0;
        }
    }

    {    # check column existing
        my $sql_query = qq{
            SHOW FIELDS
            FROM $table
            LIKE "$column"
        };
        my $sth = $dbh->prepare($sql_query);
        $sth->execute();
        my ($field) = $sth->fetchrow_array;

        if ( not $field ) {
            print " " x 4, "Column $column does not exist\n";
            return 0;
        }
    }

    {    # check values in column
        my $sql_query = qq{
            SELECT COUNT($column)
            FROM $table
        };
        my $sth = $dbh->prepare($sql_query);
        $sth->execute;
        my ($count) = $sth->fetchrow_array;

        if ( not $count ) {
            print " " x 4, "Column $column has no records\n";
        }

        return $count;
    }
}

sub quantile {
    my ( $self, $data, $part_number ) = @_;

    my $stat = Statistics::Descriptive::Full->new();

    $stat->add_data(@$data);

    my $min = $stat->min;
    my @quantiles;
    my $base = 100 / $part_number;
    for ( 1 .. $part_number - 1 ) {
        my $percentile = $stat->percentile( $_ * $base );
        push @quantiles, $percentile;
    }
    my $max = $stat->max;

    return [ $min, @quantiles, $max, ];
}

sub quantile_sql {
    my ( $self, $option, $part_number ) = @_;

    # init objects
    my $dbh = $self->dbh;

    # bind value
    my $bind_value = $option->{bind_value};
    unless ( defined $bind_value ) {
        $bind_value = [];
    }

    # init DBI query
    my $sql_query = $option->{sql_query};
    my $sth       = $dbh->prepare($sql_query);
    $sth->execute(@$bind_value);

    my @data;

    while ( my @row = $sth->fetchrow_array ) {
        push @data, $row[0];
    }

    return $self->quantile( \@data, $part_number );
}

sub calc_threshold {
    my $self = shift;

    my $dbh = $self->dbh;

    my ( $combine, $piece );

    my $sth = $dbh->prepare(
        q{
        SELECT SUM(FLOOR(align_comparables / 500) * 500)
        FROM align
        }
    );
    $sth->execute;
    my ($total_length) = $sth->fetchrow_array;

    if ( $total_length <= 5_000_000 ) {
        $piece = 10;
    }
    elsif ( $total_length <= 10_000_000 ) {
        $piece = 10;
    }
    elsif ( $total_length <= 100_000_000 ) {
        $piece = 20;
    }
    elsif ( $total_length <= 1_000_000_000 ) {
        $piece = 50;
    }
    else {
        $piece = 100;
    }

    if ( $total_length <= 1_000_000 ) {
        $combine = 100;
    }
    elsif ( $total_length <= 5_000_000 ) {
        $combine = 500;
    }
    else {
        $combine = 1000;
    }

    return ( $combine, $piece );
}

sub draw_y {
    my $self   = shift;
    my $sheet  = shift;
    my $option = shift;

    my $workbook   = $self->workbook;
    my $sheet_name = $sheet->get_name;

    my $font_name = $option->{font_name} || $self->font_name;
    my $font_size = $option->{font_size} || $self->font_size;
    my $height    = $option->{height}    || $self->height;
    my $width     = $option->{width}     || $self->width;

    # E2
    my $top  = $option->{top}  || 1;
    my $left = $option->{left} || 4;

    # 0 based
    my $first_row = $option->{first_row};
    my $last_row  = $option->{last_row};
    my $x_column  = $option->{x_column};
    my $y_column  = $option->{y_column};

    # Set axes' scale
    my $x_max_scale = $option->{x_max_scale};
    my $x_min_scale = $option->{x_min_scale};
    if ( !defined $x_min_scale ) {
        $x_min_scale = 0;
    }
    if ( !defined $x_max_scale ) {
        my $x_scale_unit = $option->{x_scale_unit};
        my $x_min_value  = min( @{ $option->{x_data} } );
        my $x_max_value  = max( @{ $option->{x_data} } );
        $x_min_scale = int( $x_min_value / $x_scale_unit ) * $x_scale_unit;
        $x_max_scale = ( int( $x_max_value / $x_scale_unit ) + 1 ) * $x_scale_unit;
    }

    my $y_scale;
    if ( exists $option->{y_data} ) {
        $y_scale = $self->_find_scale( $option->{y_data} );
    }

    my $chart = $workbook->add_chart( type => 'scatter', embedded => 1 );

    # [ $sheetname, $row_start, $row_end, $col_start, $col_end ]
    #  #"=$sheetname" . '!$A$2:$A$7',
    $chart->add_series(
        categories => [ $sheet_name, $first_row, $last_row, $x_column, $x_column ],
        values     => [ $sheet_name, $first_row, $last_row, $y_column, $y_column ],
        line       => {
            width     => 1.25,
            dash_type => 'solid',
        },
        marker => { type => 'diamond' },
    );
    $chart->set_size( width => $width, height => $height );

    # Remove title and legend
    $chart->set_title( none => 1 );
    $chart->set_legend( none => 1 );

    # Blank data is shown as a gap
    $chart->show_blanks_as('gap');

    # set axis
    $chart->set_x_axis(
        name      => $self->_replace_text( $option->{x_title} ),
        name_font => { name => $self->font_name, size => $self->font_size, },
        num_font  => { name => $self->font_name, size => $self->font_size, },
        line            => { color   => 'black', },
        major_gridlines => { visible => 0, },
        minor_gridlines => { visible => 0, },
        min             => $x_min_scale,
        max             => $x_max_scale,
    );
    $chart->set_y_axis(
        name      => $self->_replace_text( $option->{y_title} ),
        name_font => { name => $self->font_name, size => $self->font_size, },
        num_font  => { name => $self->font_name, size => $self->font_size, },
        line            => { color   => 'black', },
        major_gridlines => { visible => 0, },
        minor_gridlines => { visible => 0, },
        defined $y_scale
        ? ( min => $y_scale->{min}, max => $y_scale->{max}, major_unit => $y_scale->{unit}, )
        : (),
    );

    # plorarea
    $chart->set_plotarea( border => { color => 'black', }, );

    $sheet->insert_chart( $top, $left, $chart );

    return;
}

sub draw_2y {
    my $self   = shift;
    my $sheet  = shift;
    my $option = shift;

    my $workbook   = $self->workbook;
    my $sheet_name = $sheet->get_name;

    my $font_name = $option->{font_name} || $self->font_name;
    my $font_size = $option->{font_size} || $self->font_size;
    my $height    = $option->{height}    || $self->height;
    my $width     = $option->{width}     || $self->width;

    # E2
    my $top  = $option->{top}  || 1;
    my $left = $option->{left} || 4;

    # 0 based
    my $first_row = $option->{first_row};
    my $last_row  = $option->{last_row};
    my $x_column  = $option->{x_column};
    my $y_column  = $option->{y_column};
    my $y2_column = $option->{y2_column};

    # Set axes' scale
    my $x_max_scale = $option->{x_max_scale};
    my $x_min_scale = $option->{x_min_scale};
    if ( !defined $x_min_scale ) {
        $x_min_scale = 0;
    }
    if ( !defined $x_max_scale ) {
        my $x_scale_unit = $option->{x_scale_unit};
        my $x_min_value  = min( @{ $option->{x_data} } );
        my $x_max_value  = max( @{ $option->{x_data} } );
        $x_min_scale = int( $x_min_value / $x_scale_unit ) * $x_scale_unit;
        $x_max_scale = ( int( $x_max_value / $x_scale_unit ) + 1 ) * $x_scale_unit;
    }

    my $y_scale;
    if ( exists $option->{y_data} ) {
        $y_scale = $self->_find_scale( $option->{y_data} );
    }

    my $y2_scale;
    if ( exists $option->{y2_data} ) {
        $y2_scale = $self->_find_scale( $option->{y2_data} );
    }

    my $chart = $workbook->add_chart( type => 'scatter', embedded => 1 );

    # [ $sheetname, $row_start, $row_end, $col_start, $col_end ]
    #  #"=$sheetname" . '!$A$2:$A$7',
    $chart->add_series(
        categories => [ $sheet_name, $first_row, $last_row, $x_column, $x_column ],
        values     => [ $sheet_name, $first_row, $last_row, $y_column, $y_column ],
        line       => {
            width     => 1.25,
            dash_type => 'solid',
        },
        marker => { type => 'diamond', },
    );

    # second Y axis
    $chart->add_series(
        categories => [ $sheet_name, $first_row, $last_row, $x_column,  $x_column ],
        values     => [ $sheet_name, $first_row, $last_row, $y2_column, $y2_column ],
        line       => {
            width     => 1.25,
            dash_type => 'solid',
        },
        marker  => { type => 'square', size => 5, fill => { color => 'white', }, },
        y2_axis => 1,
    );
    $chart->set_size( width => $width, height => $height );

    # Remove title and legend
    $chart->set_title( none => 1 );
    $chart->set_legend( none => 1 );

    # Blank data is shown as a gap
    $chart->show_blanks_as('gap');

    # set axis
    $chart->set_x_axis(
        name      => $self->_replace_text( $option->{x_title} ),
        name_font => { name => $self->font_name, size => $self->font_size, },
        num_font  => { name => $self->font_name, size => $self->font_size, },
        line            => { color   => 'black', },
        major_gridlines => { visible => 0, },
        minor_gridlines => { visible => 0, },
        min             => $x_min_scale,
        max             => $x_max_scale,
    );
    $chart->set_y_axis(
        name      => $self->_replace_text( $option->{y_title} ),
        name_font => { name => $self->font_name, size => $self->font_size, },
        num_font  => { name => $self->font_name, size => $self->font_size, },
        line            => { color   => 'black', },
        major_gridlines => { visible => 0, },
        minor_gridlines => { visible => 0, },
        defined $y_scale
        ? ( min => $y_scale->{min}, max => $y_scale->{max}, major_unit => $y_scale->{unit}, )
        : (),
    );
    $chart->set_y2_axis(
        name      => $self->_replace_text( $option->{y2_title} ),
        name_font => { name => $self->font_name, size => $self->font_size, },
        num_font  => { name => $self->font_name, size => $self->font_size, },
        line            => { color   => 'black', },
        major_gridlines => { visible => 0, },
        minor_gridlines => { visible => 0, },
        defined $y2_scale
        ? ( min => $y2_scale->{min}, max => $y2_scale->{max}, major_unit => $y2_scale->{unit}, )
        : (),
    );

    # plorarea
    $chart->set_plotarea( border => { color => 'black', }, );

    $sheet->insert_chart( $top, $left, $chart );

    return;
}

sub _find_scale {
    my $self    = shift;
    my $dataset = shift;

    my $axis = Chart::Math::Axis->new;

    $axis->add_data( @{$dataset} );
    $axis->set_maximum_intervals( $self->max_ticks );

    return {
        max  => $axis->top,
        min  => $axis->bottom,
        unit => $axis->interval_size,
    };
}

sub _replace_text {
    my $self    = shift;
    my $text    = shift;
    my $replace = $self->replace;

    for my $key ( keys %$replace ) {
        my $value = $replace->{$key};
        $text =~ s/$key/$value/gi;
    }

    return $text;
}

# instance destructor
# invoked only as object method
sub DESTROY {
    my ($self) = shift;

    # close excel objects
    my $workbook = $self->workbook;
    $workbook->close if $workbook;

    if ( !$self->mocking ) {

        # close dbh
        my $dbh = $self->dbh;
        $dbh->disconnect if $dbh;
    }
}

1;

__END__

=head1 NAME

AlignDB::ToXLSX - create xlsx files

=head1 SYNOPSIS

    my $write_obj = AlignDB::ToXLSX->new(
        outfile => $outfile,
        mocking => 1,
    );

=cut

=head1 LICENSE

Copyright 2014- Qiang Wang

This library is free software; you can redistribute it and/or modify it under the same terms as Perl itself.

=head1 AUTHOR

Qiang Wang

=cut

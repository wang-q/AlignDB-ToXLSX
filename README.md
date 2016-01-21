# NAME

AlignDB::ToXLSX - Create xlsx files from arrays or SQL queries.

# SYNOPSIS

    # Mysql
    my $write_obj = AlignDB::ToXLSX->new(
        outfile => $outfile,
        dbh     => $dbh,
    );
    
    # MongoDB
    my $write_obj = AlignDB::ToXLSX->new(
        outfile => $outfile,
    );
    

# AUTHOR

Qiang Wang &lt;wang-q@outlook.com>

# COPYRIGHT AND LICENSE

This software is copyright (c) 2008 by Qiang Wang.

This is free software; you can redistribute it and/or modify it under
the same terms as the Perl 5 programming language system itself.

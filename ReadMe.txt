This project provides a VB.Net class to parse the text of an address into its component parts.  This is useful for standardizing addresses and for matching them to each other.  It is also the first step in geocoding an address.

The parsing routine was developed through extensive testing using addresses reported to the San Francisco Department of Public Health for various reasons.  There were over 800,000 addresses used to test the logic, and less than 500 addresses that began with a number could not be parsed.  And very few did not appear to parse the way a person would read them.

It is extremely long and complex to handle all the variations reported to us.  While it was developed for San Francisco addresses, the parsing logic applies to addresses throughout the United States.  

Components
-----------

Note that the only components that are always present are house number and street name.  The other components are zero-byte strings if they are not found in the address text.

The components returned are the following:

1. House number 

The number for the address.  If there is no number for the address, it will not parse.  Values are always numeric, and do not include any leading zeros.  However, the property itself is a string, not a number, so that it will not create an error if the value is not numeric.

2. House number suffix

This is the letter following the number (if any).  "123A Main Street" has 123 for the number and "A" for the suffix.  This allows us to treat the house number as a numeric value without losing this information.

3. Predirectional

This is the compass direction that may precede the street name.  Taking it out of the street name helps standardize the address; "South Van Ness" is really the same as "S Van Ness" and "So Van Ness".

Values are restricted to the four cardinal directions (north, south, east, and west) and the four ordinal directions (northeast, southeast, southwest, and northwest) and are represented by one letter for cardinal directions and two letters for ordinal directions.

4. Street name

This is the name of the street.  Note that the program does not validate that the street exists, but it fixes some common typing errors that prevent the program from finding the other parts of the address.

Note that the street name is the only part of the address where the number of words in it is not fixed.  The logic of the parser is essentially to find all the other parts and count whatever is left as the street name.

5. Street type

This is type after the name, e.g., "St", "Ave", "Blvd".  Values are restricted to the abbreviations for the approved USPS street suffixes (see https://pe.usps.com/text/pub28/28apc_002.htm).

6. Postdirectional

This is the compass direction that may follow the street name, e.g., "123 Buena Vista Avenue West".  Like the predirectional, taking it out of the street name helps standardize the address, and values are restricted to the four cardinal directions and the four ordinal directions.

7. Secondary unit designators (SUD) 

Many address have type of unit after the street address, e.g., "Apartment 101", "Suite 404", "2nd Floor".  This is the SUD.

SUDs can have a type and a range.  The type usually comes first (e.g., "Apt", "Unit", "Rm").  Values for type are restricted to the abbreviations for the approved USPS secondary unit designators (see https://pe.usps.com/text/pub28/pub28apc_003.htm).

The range is the identifier after the type, e.g., the "101" in "Unit 101".  They can be numbers, letters, or combinations.

Note that some SUD types do not have a range, e.g., "Lower", "Rear", "Basement".  And an address may have a range without the SUD identified, e.g., "123 Main St #101" or "321 1st St B".  But each SUD will have one or the other if not both.

Success codes
-------------

In addition to the components, the class exposes a number indicating whether the address text was successfully parsed or why it was not.

A value of zero indicates that the address was successfully parsed, a value of (-1) indicates that the method to parse the input address was not called.  The remaining values (all greater than zero) indicate that the address text could not be parsed.

1 = NO_NUMBER 

Address does not start with a digit from 1 to 9 (after removing any leading zeros).

2 = ONLY_NUMBER 

The address text is only a number (or number and letter).

3 = PO_BOX 

The address is a post office box.

4 = INTERSECTION 

Address was a street intersection without any valid street address.

5 = NO_STREET 

No street name could be found (e.g., "123 St Apt 4").

6 = TWO_NUMBERS 

More than one house number was found at the beginning.

9 = UNKNOWN_ERROR 

Any other error.

Parser class
------------

The Parser class has a Parse method to parse the address specified and expose each of the components and the success code as properties.  The constructor allows the address text to be passed and the parsing to be done without calling the Parse method explicitly.  

The RawText property exposed the address text passed to the parser.  Setting this property will also invoke the parser.

An address may have more than one SUD specified, e.g., "123 Main St Building 100 Room A".  For this reason, the parser class exposes SUDs as a collection, with each SUD specifying the type and/or the range in the address.  An address without any SUD will still have that collection attached to it, but the count of SUDs for the address will be zero.

The parser makes extensive use of regular expressions.  It keeps them in a list of RegEx objects as they are created so they can be reused.  There is a constructor that allows you to compile these as they are created, which may improve performance when parsing a long list of addresses.

Multi-line addresses
--------------------

If the text of the adddress includes a line break or a slash, the parser will attempt to parse the beginning part of the address.  If it cannot be parsed, it will proceed to try to parse each subsequent part, up to three times.

"Somename Hotel / 123 Main St" will therfore parse as "123 Main St".

Special cases
-------------

The parsing routine handles a number of addresses that are do not follow the usual standard:

1. Avenues

For addresses like "123 Avenue A" the parser returns "A" as the street name and "Ave" as the street type.

2. Floors

For addresses like "123 Main St 2nd Floor" the parser returns "2" as the range and "FL" as the SUD type, even though these are reversed from the usual order.

3. One-half

Addresses like "123 1/2 Main St" save the "1/2" as the street suffix, leaving the house number as a whole number.

4. Foreign language street types

In many foreign languages, the street type precedes the street name.  The routine can handle a few of them, including "Avenida", "Calle", and "Rue".  It save the street type in its English translation.

/**
 * utilities that work with xlsx library
 * 
 * pass the library in as the argument
 */

module.exports = function(XLSX) {

    /**
     * inverts the result of a predicate function
     */
    const not = predicate => v => !predicate(v);
    
    
    /**
     * boolean and of two predicates
     */
    const and = (p1, p2) => v => p1(v) && p2(v);
    
    
    /**
     * boolean or of two predicates
     */
    const or = (p1, p2) => v => p1(v) || p2(v);
    
    
    /**
     * checks whether a sheet key is a special reference
     * @param {string} ref 
     */
    const isSpecialRef = ref => ref[0] === '!';
    
    
    /**
     * checks whether a ref is a range reference
     * @param {string} ref 
     */
    const isRangeRef = ref => ref.indexOf(':') !== -1;
    
    
    /**
     * checks whether a sheet key is a cell reference
     * @param {string} ref 
     */
    const isCellRef = not(or(isSpecialRef, isRangeRef));
    
    
    /**
     * gets a sheet from a workbook by name, also splices in the name under the '!name' key in the sheet
     * @param {WorkBook} sheet 
     * @param {string} name 
     */
    const getSheet = workbook => name => Object.assign(workbook.Sheets[name], { '!name': name });
    
    
    /**
     * gets the range of a sheet
     * @param {WorkSheet} sheet 
     */
    const getSheetRange = sheet => sheet['!ref'];
    
    
    /**
     * returns all cell references of a sheet
     * @param {WorkSheet} sheet 
     */
    const getCellRefs = sheet => Object.keys(sheet).filter(isCellRef);
    
    
    /**
     * gets the value of a cell
     * @param {WorkSheet} sheet 
     * @param {string} ref 
     */
    const getCellValue = sheet => ref => sheet[ref].v;
    

    /**
     * sets the value of a cell
     * @param {WorkSheet} sheet
     * @param {string} ref
     * @param {any} value
     */
    const setCellValue = sheet => ref => v => {
        var t = 'z';
        var z = "";
        if (typeof v == 'number') t = 'n';
        else if (typeof v == 'boolean') t = 'b';
        else if (typeof v == 'string') t = 's';
        else if (v instanceof Date) {
            t = 'd';
            // if (!o.cellDates) { t = 'n'; v = datenum(v); }
            // z = o.dateNF || SSF._table[14];
        }
        sheet[ref] = cell = ({ t, v });
        // if (z) cell.z = z;
    }

    
    /**
     * gets the cr address of a cell, alias of XLSX.utils.decode_cell;
     * @param {string} ref
     */
    const getCellAddress = XLSX.utils.decode_cell;
    
    
    /**
     * gets the cr address of a range, alias of XLSX.utils.decode_range
     * @param {string} ref
     */
    const getRangeAddress = XLSX.utils.decode_range;
    
    
    /**
     * gets the A1 reference of a cell, alias of XLSX.utils.encode_cell
     * @param {object} address
     */
    const getCellRef = XLSX.utils.encode_cell;
    
    
    /**
     * gets the A1:B1 reference of a range, alias of XLSX.utils.encode_range
     * @param {object} address
     */
    const getRangeRef = XLSX.utils.encode_range;
    
    
    /**
     * gets the width of the range or infinity if it is a row range
     * @param {string} ref range reference
     */
    const getRangeWidth = ref => {
        if (isCellRef(ref)) return 1;
        let r = getRangeAddress(ref);
        return (r.s.c < 0) ? Number.POSITIVE_INFINITY : (1 + r.e.c - r.s.c);
    }
    
    
    /**
     * gets the height of the range or infinity if it is a column range
     * @param {string} ref range reference
     */
    const getRangeHeight = ref => {
        if (isCellRef(ref)) return 1;
        let r = getRangeAddress(ref);
        return (r.s.r < 0) ? Number.POSITIVE_INFINITY : (1 + r.e.r - r.s.r);
    }
    
    
    /**
     * gets the range width and height
     * @param {string} ref range reference
     */
    const getRangeSize = ref => ({w: getRangeWidth(ref), h: getRangeHeight(ref)});
    
    
    /**
     * get all the cell refs of the sheet that are in the range
     * @param {WorkSheet} sheet 
     * @param {string} range 
     */
    const getRangeCellRefs = (sheet, range) => getCellRefs(sheet).filter(inRange(range));
    
    
    /**
     * gets all names ranges from the workbook
     * @param {WorkBook} workbook 
     */
    const getNames = workbook => workbook.Workbook.Names;
    
    
    /**
     * checks whether a cell is in a range
     * @param {string} range 
     * @param {string} ref
     */
    const inRange = range => ref => {
        let a = getCellAddress(ref);
        let r = getRangeAddress(range);
        const within = (s, v, e) => (s <= v) && (v <= e);
        return (
            ((r.s.c < 0) || within(r.s.c, a.c, r.e.c)) &&
            ((r.s.r < 0) || within(r.s.r, a.r, r.e.r))
        )
    }
    
    
    /**
     * updates the sheet range to reflect the current cells
     * @param {WorkSheet} sheet 
     */
    const updateRange = sheet => {
        let addresses = getCellRefs(sheet).map(getCellAddress);
        if (!addresses.length) return;
        let minCol = addresses.reduce((min, { c, r }) => Math.min(min, c), Number.POSITIVE_INFINITY);
        let maxCol = addresses.reduce((max, { c, r }) => Math.max(max, c), Number.NEGATIVE_INFINITY);
        let minRow = addresses.reduce((min, { c, r }) => Math.min(min, r), Number.POSITIVE_INFINITY);
        let maxRow = addresses.reduce((max, { c, r }) => Math.max(max, r), Number.NEGATIVE_INFINITY);
        let minRef = getCellRef({ c: minCol, r: minRow });
        let maxRef = getCellRef({ c: maxCol, r: maxRow });
        sheet['!ref'] = `${minRef}:${maxRef}`;
        return sheet;
    }
    
    
    /**
     * checks whether a cell is above a certain row (0 based)
     * @param {number} row 
     * @param {string} ref 
     */
    const isAbove = row => ref => getCellAddress(ref).r < row;
    
    
    /**
     * checks whether a cell is on a certain row (0 based)
     * @param {number} row 
     * @param {string} ref 
     */
    const isAtRow = row => ref => getCellAddress(ref).r == row;
    
    
    /**
     * checks whether a cell is below a certain row (0 based)
     * @param {number} row 
     * @param {string} ref 
     */
    const isBelow = row => ref => getCellAddress(ref).r > row;
    
    
    /**
     * checks whether a cell is before a certain column (0 based)
     * @param {number} col 
     * @param {string} ref 
     */
    const isBefore = col => ref => getCellAddress(ref).c < col;
    
    
    /**
     * checks whether a cell is on a certain column (0 based)
     * @param {number} col 
     * @param {string} ref 
     */
    const isAtCol = col => ref => getCellAddress(ref).c == col;
    
    
    /**
     * checks whether a cell is after a certain column (0 based)
     * @param {number} col 
     * @param {string} ref 
     */
    const isAfter = col => ref => getCellAddress(ref).c > col;
    

    /**
     * creates a sorting function (to be used in array.sort) that sorts
     * on a particular dimension of a cell (row or column), optionally in reverse
     * @param {string} axis to sort, 'c' or 'r'
     * @param {boolean} reverse sort in reverse, defaults to false
     */    
    const sortDim = (axis, reverse = false) => (refa, refb) => {
        let aa = getCellAddress(refa);
        let ab = getCellAddress(refb);
        if (aa[axis] === ab[axis]) return 0;
        return ((aa[axis] < ab[axis]) !== reverse) ? -1 : 1;
    }
    

    /**
     * creates a sorting function to sorts cells in the row direction
     * @param {boolean} reverse 
     */
    const sortRow = reverse => sortDim('r', reverse);


    /**
     * creates a sorting function to sorts cells in the column direction
     * @param {boolean} reverse 
     */
    const sortCol = reverse => sortDim('c', reverse);


    /**
     * creates a sorting function to sorts cells from bottom to top
     */
    const sortBottomTop = sortRow(true);
    
    
    /**
     * creates a sorting function to sorts cells from top to bottom
     */
    const sortTopBottom = sortRow(false);


    /**
     * creates a sorting function to sorts cells from end to start (right to left in ltr writing)
     */
    const sortEndStart = sortCol(true);


    /**
     * creates a sorting function to sorts cells from start to end (left to right in ltr writing)
     */
    const sortStartEnd = sortCol(false);
    
    
    /**
     * returns a cell ref relative to the given ref by the given delta
     * @param {number[]} delta relative offset in column and row direction 
     * @param {string} ref cell reference
     */
    const relCell = ([dc, dr]) => (ref) => {
        let { c, r } = getCellAddress(ref);
        return getCellRef({ c: c + dc, r: r + dr });
    }
    
    
    /**
     * returns a range ref relative to the given ref by the given delta
     * @param {number[]} delta relative offset in column and row direction
     * @param {string} ref range reference 
     */
    const relRange = ([dc, dr]) => ref => {
        let r = getRangeAddress(ref);
        if (r.s.c >= 0) { r.s.c += dc; r.e.c += dc; }
        if (r.s.r >= 0) { r.s.r += dr; r.e.r += dr; }
        return getRangeRef(r);
    }
    
    
    /**
     * returns a range or cell ref relative to the given ref by the given delta
     * @param {number} delta 
     * @param {string} ref cell or range ref
     */
    const rel = delta => ref => {
        switch (true) {
            case isRangeRef(ref): return relRange(delta)(ref);
            case isCellRef(ref): return relCell(delta)(ref);
            default: return ref;
        }
    }
    
    
    /**
     * moves the cell content from a location to a new location
     * it overwrites anything that is in the new location
     * and leaves the old location blank
     * NOTE: it does not update formulas
     * @param {WorkSheet} sheet 
     * @param {string} to cell reference
     * @param {string} from cell reference
     */
    const moveCell = (sheet, to) => (from) => {
        sheet[to] = sheet[from];
        delete sheet[from];
        //todo: update formulae
    }
    
    const moveCellBy = (sheet, delta) => (ref) => moveCell(sheet, relCell(delta)(ref))(ref);
    
    const insertRow = (sheet) => (index) => {
        getCellRefs(sheet)
        .sort(sortBottomTop)
        .filter(or(isBelow(index), isAtRow(index)))
        .forEach(moveCellBy(sheet, [0, 1]));
        return updateRange(sheet);
        //todo: update merges and styles and row heights
    }
    
    const insertColumn = (sheet) => (index) => {
        getCellRefs(sheet)
        .sort(sortEndStart)
        .filter(or(isAfter(index), isAtCol(index)))
        .forEach(moveCellBy(sheet, [1, 0]));
        return updateRange(sheet);
        //todo: update merges and styles and row heights
    }
    
    const insertCellShiftDown = (sheet) => (ref) => {
        let { c, r } = getCellAddress(ref);
        getCellRefs(sheet)
            .sort(sortBottomTop)
            .filter(or(isBelow(r), isAtRow(r)))
            .filter(isAtCol(c))
            .forEach(moveCellBy(sheet, [0, 1]));
        return updateRange(sheet);
    }
    const insertCellShiftEnd = (sheet) => (ref) => {
        let { c, r } = getCellAddress(ref);
        getCellRefs(sheet)
            .sort(sortEndStart)
            .filter(or(isAfter(c), isAtCol(c)))
            .filter(isAtRow(r))
            .forEach(moveCellBy(sheet, [1, 0]));
        return updateRange(sheet);
    }
    const copyRangeDown = (sheet) => (range) => {
    
    }
    const copyRangeEnd = (sheet) => (range) => {
    
    }
    
    const deleteColumn = (sheet) => (index) => {
        getCellRefs(sheet)
            .filter(isAtCol(index))
            .forEach(clearCell(sheet));
        getCellRefs(sheet)
            .sort(sortStartEnd)
            .filter(isAfter(index))
            .forEach(moveCellBy(sheet, [-1, 0]));
        return updateRange(sheet);
    }
    
    const deleteRow = (sheet) => (index) => {
        getCellRefs(sheet)
            .filter(isAtRow(index))
            .forEach(clearCell(sheet));
        getCellRefs(sheet)
            .sort(sortTopBottom)
            .filter(isBelow(index))
            .forEach(moveCellBy(sheet, [0, -1]));
        return updateRange(sheet);
    }
    
    const clearCell = (sheet) => (ref) => {
        delete sheet[ref];
        return sheet;
    }
    
    /**
     * checks whether a parent range fully contains a child range
     * @param {WorkSheet} sheet 
     * @param {string} parentRange 
     * @param {string} childRange 
     */
    const contains = (sheet) => (parentRange) => (childRange) => {
        return getRangeCellRefs(sheet, childRange)
            .every(inRange(parentRange));
    }
    
    
    /**
     * checks whether a parent range partly contains a child range
     * @param {WorkSheet} sheet 
     * @param {string} parentRange 
     * @param {string} childRange 
     */
    const overlaps = (sheet) => (parentRange) => (childRange) => {
        return getRangeCellRefs(sheet, childRange)
            .some(inRange(parentRange));
    }
    
    return {
        not,
        isSpecialRef,
        isCellRef,
        getSheet,
        getSheetRange,
        getCellRefs,
        getCellValue,
        setCellValue,
        getCellAddress,
        getCellRef,
        getRangeAddress,
        getRangeRef,
        getRangeWidth,
        getRangeHeight,
        getRangeSize,
        getRangeCellRefs,
        getNames,
        inRange,
        updateRange,
        isAbove,
        isAtRow,
        isBelow,
        isBefore,
        isAtCol,
        isAfter,
        sortDim,
        sortRow,
        sortCol,
        sortBottomTop,
        sortTopBottom,
        sortEndStart,
        sortStartEnd,
        relCell,
        relRange,
        rel,
        moveCell,
        moveCellBy,
        insertRow,
        insertColumn,
        insertCellShiftDown,
        insertCellShiftEnd,
        copyRangeDown,
        copyRangeEnd,
        deleteColumn,
        deleteRow,
        clearCell,
        contains,
        overlaps
    }
}


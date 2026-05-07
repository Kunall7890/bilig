export function isWasmSafeBuiltinArity(callee: string, argc: number): boolean {
  switch (callee) {
    case 'NOT':
    case 'LEN':
    case 'YEAR':
    case 'MONTH':
    case 'DAY':
    case 'HOUR':
    case 'MINUTE':
    case 'SECOND':
    case 'INT':
    case 'SIN':
    case 'COS':
    case 'TAN':
    case 'ASIN':
    case 'ACOS':
    case 'ATAN':
    case 'DEGREES':
    case 'RADIANS':
    case 'EXP':
    case 'LN':
    case 'LOG10':
    case 'SQRT':
      return argc === 1
    case 'TODAY':
    case 'NOW':
    case 'RAND':
      return argc === 0
    case 'NA':
      return argc === 0
    case 'IF':
      return argc === 3
    case 'IFS':
      return argc >= 2 && argc % 2 === 0
    case 'IFERROR':
    case 'IFNA':
      return argc === 2
    case 'WEEKDAY':
      return argc === 1 || argc === 2
    case 'DAYS':
      return argc === 2
    case 'COUNTBLANK':
      return argc >= 1
    case 'CHOOSE':
      return argc >= 2
    case 'DAYS360':
    case 'YEARFRAC':
      return argc === 2 || argc === 3
    case 'DISC':
    case 'INTRATE':
    case 'RECEIVED':
    case 'PRICEDISC':
    case 'YIELDDISC':
      return argc === 4 || argc === 5
    case 'COUPDAYBS':
    case 'COUPDAYS':
    case 'COUPDAYSNC':
    case 'COUPNCD':
    case 'COUPNUM':
    case 'COUPPCD':
      return argc === 3 || argc === 4
    case 'PRICEMAT':
    case 'YIELDMAT':
    case 'DURATION':
    case 'MDURATION':
      return argc === 5 || argc === 6
    case 'ODDFPRICE':
    case 'ODDFYIELD':
    case 'ODDLPRICE':
    case 'ODDLYIELD':
      return argc === 7 || argc === 8
    case 'TBILLPRICE':
    case 'TBILLYIELD':
    case 'TBILLEQ':
      return argc === 3
    case 'IRR':
      return argc === 1 || argc === 2
    case 'MIRR':
      return argc === 3
    case 'XNPV':
      return argc === 3
    case 'XIRR':
      return argc === 2 || argc === 3
    case 'PRICE':
    case 'YIELD':
      return argc === 6 || argc === 7
    case 'ISOWEEKNUM':
    case 'TIMEVALUE':
      return argc === 1
    case 'WEEKNUM':
      return argc === 1 || argc === 2
    case 'WORKDAY':
    case 'NETWORKDAYS':
      return argc === 2 || argc === 3
    case 'WORKDAY.INTL':
    case 'NETWORKDAYS.INTL':
      return argc >= 2 && argc <= 4
    case 'COUNTIF':
    case 'USE.THE.COUNTIF':
      return argc === 2
    case 'COUNTIFS':
      return argc >= 2 && argc % 2 === 0
    case 'DAVERAGE':
    case 'DCOUNT':
    case 'DCOUNTA':
    case 'DGET':
    case 'DMAX':
    case 'DMIN':
    case 'DPRODUCT':
    case 'DSTDEV':
    case 'DSTDEVP':
    case 'DSUM':
    case 'DVAR':
    case 'DVARP':
      return argc === 3
    case 'ADDRESS':
      return argc >= 2 && argc <= 5
    case 'SUMIF':
    case 'AVERAGEIF':
      return argc === 2 || argc === 3
    case 'SUMIFS':
    case 'AVERAGEIFS':
      return argc >= 3 && argc % 2 === 1
    case 'REPLACE':
      return argc === 4
    case 'SUBSTITUTE':
      return argc === 3 || argc === 4
    case 'REPT':
      return argc === 2
    case 'TEXT':
      return argc === 2
    case 'PHONETIC':
      return argc === 1
    case 'TEXTBEFORE':
    case 'TEXTAFTER':
      return argc >= 2 && argc <= 6
    case 'TEXTSPLIT':
      return argc >= 2 && argc <= 6
    case 'TEXTJOIN':
      return argc >= 3
    case 'POWER':
    case 'CONVERT':
      return argc === 3
    case 'EXACT':
    case 'ATAN2':
      return argc === 2
    case 'BESSELI':
    case 'BESSELJ':
    case 'BESSELK':
    case 'BESSELY':
      return argc === 2
    case 'EUROCONVERT':
      return argc >= 3 && argc <= 5
    case 'UPPER':
    case 'LOWER':
    case 'TRIM':
    case 'VALUE':
    case 'CHAR':
    case 'CODE':
    case 'UNICODE':
    case 'UNICHAR':
    case 'CLEAN':
    case 'ASC':
    case 'JIS':
    case 'DBCS':
    case 'BAHTTEXT':
    case 'LENB':
    case 'SINH':
    case 'COSH':
    case 'TANH':
    case 'ASINH':
    case 'ACOSH':
    case 'ATANH':
    case 'ACOT':
    case 'ACOTH':
    case 'COT':
    case 'COTH':
    case 'CSC':
    case 'CSCH':
    case 'SEC':
    case 'SECH':
    case 'SIGN':
    case 'EVEN':
    case 'ODD':
    case 'FACT':
    case 'FACTDOUBLE':
      return argc === 1
    case 'NUMBERVALUE':
      return argc >= 1 && argc <= 3
    case 'VALUETOTEXT':
      return argc === 1 || argc === 2
    case 'DOLLAR':
      return argc >= 1 && argc <= 3
    case 'DOLLARDE':
    case 'DOLLARFR':
    case 'COMBIN':
    case 'COMBINA':
    case 'QUOTIENT':
      return argc === 2
    case 'BASE':
      return argc === 2 || argc === 3
    case 'DECIMAL':
      return argc === 2
    case 'BIN2DEC':
    case 'HEX2DEC':
    case 'OCT2DEC':
      return argc === 1
    case 'BIN2HEX':
    case 'BIN2OCT':
    case 'DEC2BIN':
    case 'DEC2HEX':
    case 'DEC2OCT':
    case 'HEX2BIN':
    case 'HEX2OCT':
    case 'OCT2BIN':
    case 'OCT2HEX':
      return argc === 1 || argc === 2
    case 'BITAND':
    case 'BITOR':
    case 'BITXOR':
      return argc >= 2
    case 'BITLSHIFT':
    case 'BITRSHIFT':
      return argc === 2
    case 'MATCH':
      return argc === 2 || argc === 3
    case 'CORREL':
    case 'COVAR':
    case 'PEARSON':
    case 'COVARIANCE.P':
    case 'COVARIANCE.S':
    case 'PERCENTRANK':
    case 'PERCENTRANK.INC':
    case 'PERCENTRANK.EXC':
    case 'SMALL':
    case 'LARGE':
    case 'PERCENTILE':
    case 'PERCENTILE.INC':
    case 'PERCENTILE.EXC':
    case 'QUARTILE':
    case 'QUARTILE.INC':
    case 'QUARTILE.EXC':
    case 'RANK':
    case 'RANK.EQ':
    case 'RANK.AVG':
    case 'INTERCEPT':
    case 'RSQ':
    case 'SLOPE':
    case 'STEYX':
      return argc === 2 || argc === 3
    case 'MEDIAN':
    case 'MODE.MULT':
    case 'GCD':
    case 'LCM':
    case 'PRODUCT':
    case 'GEOMEAN':
    case 'HARMEAN':
    case 'SUMSQ':
      return argc >= 1
    case 'FREQUENCY':
      return argc === 2
    case 'PROB':
      return argc === 3 || argc === 4
    case 'TRIMMEAN':
      return argc === 2
    case 'FORECAST':
    case 'FORECAST.LINEAR':
      return argc === 3
    case 'TREND':
    case 'GROWTH':
    case 'LINEST':
    case 'LOGEST':
      return argc >= 1 && argc <= 4
    case 'XMATCH':
      return argc >= 2 && argc <= 4
    case 'XLOOKUP':
      return argc >= 3 && argc <= 6
    case 'INDEX':
      return argc === 2 || argc === 3
    case 'VLOOKUP':
    case 'HLOOKUP':
      return argc === 3 || argc === 4
    case 'LEFT':
    case 'RIGHT':
    case 'LEFTB':
    case 'RIGHTB':
      return argc === 1 || argc === 2
    case 'MID':
    case 'MIDB':
      return argc === 3
    case 'FIND':
    case 'SEARCH':
    case 'FINDB':
    case 'SEARCHB':
      return argc === 2 || argc === 3
    case 'REPLACEB':
      return argc === 4
    case 'ISBLANK':
    case 'ISNUMBER':
    case 'ISTEXT':
      return argc === 0 || argc === 1
    case 'ROUND':
    case 'ROUNDUP':
    case 'ROUNDDOWN':
    case 'FLOOR':
    case 'CEILING':
    case 'LOG':
      return argc === 1 || argc === 2
    case 'T':
    case 'N':
    case 'TYPE':
    case 'GAUSS':
    case 'PHI':
    case 'NORMSDIST':
    case 'NORMSINV':
      return argc === 1 || (argc === 0 && (callee === 'T' || callee === 'N' || callee === 'TYPE'))
    case 'DELTA':
    case 'GESTEP':
    case 'LOGNORMDIST':
    case 'EFFECT':
    case 'NOMINAL':
    case 'RRI':
    case 'PERMUT':
    case 'PERMUTATIONA':
      return argc === 2
    case 'STANDARDIZE':
    case 'NORMINV':
    case 'LOGINV':
    case 'PDURATION':
    case 'CONFIDENCE.NORM':
    case 'CONFIDENCE':
    case 'CONFIDENCE.T':
    case 'CRITBINOM':
    case 'BINOM.INV':
      return argc === 3
    case 'ERF':
      return argc === 1 || argc === 2
    case 'ERF.PRECISE':
    case 'ERFC':
    case 'ERFC.PRECISE':
    case 'FISHER':
    case 'FISHERINV':
    case 'GAMMALN':
    case 'GAMMALN.PRECISE':
    case 'GAMMA':
      return argc === 1
    case 'GAMMA.INV':
    case 'GAMMAINV':
      return argc === 3
    case 'CHIDIST':
    case 'LEGACY.CHIDIST':
    case 'CHIINV':
    case 'CHISQ.DIST.RT':
    case 'CHISQ.INV.RT':
    case 'CHISQDIST':
    case 'CHISQINV':
    case 'LEGACY.CHIINV':
    case 'CHISQ.TEST':
    case 'CHITEST':
    case 'LEGACY.CHITEST':
    case 'F.TEST':
    case 'FTEST':
      return argc === 2
    case 'Z.TEST':
    case 'ZTEST':
      return argc === 2 || argc === 3
    case 'F.DIST.RT':
    case 'FDIST':
    case 'LEGACY.FDIST':
      return argc === 3
    case 'CHISQ.INV':
      return argc === 2
    case 'CHISQ.DIST':
      return argc === 3
    case 'BETA.INV':
    case 'BETAINV':
      return argc >= 3 && argc <= 5
    case 'BETA.DIST':
      return argc >= 4 && argc <= 6
    case 'BETADIST':
      return argc >= 3 && argc <= 5
    case 'F.DIST':
      return argc === 4
    case 'T.DIST':
      return argc === 3
    case 'T.DIST.RT':
    case 'T.DIST.2T':
    case 'T.INV':
    case 'T.INV.2T':
    case 'TINV':
      return argc === 2
    case 'TDIST':
      return argc === 3
    case 'T.TEST':
    case 'TTEST':
      return argc === 4
    case 'F.INV':
    case 'F.INV.RT':
    case 'FINV':
    case 'LEGACY.FINV':
      return argc === 3
    case 'WEIBULL':
    case 'WEIBULL.DIST':
    case 'GAMMADIST':
    case 'GAMMA.DIST':
    case 'BINOMDIST':
    case 'BINOM.DIST':
    case 'NEGBINOM.DIST':
      return argc === 4
    case 'EXPONDIST':
    case 'EXPON.DIST':
    case 'POISSON':
    case 'POISSON.DIST':
    case 'NEGBINOMDIST':
      return argc === 3
    case 'BINOM.DIST.RANGE':
      return argc === 3 || argc === 4
    case 'HYPGEOMDIST':
      return argc === 4
    case 'HYPGEOM.DIST':
      return argc === 5
    case 'NORMDIST':
      return argc === 4
    case 'NORM.DIST':
      return argc === 4
    case 'NORM.INV':
      return argc === 3
    case 'NORM.S.DIST':
      return argc === 1 || argc === 2
    case 'NORM.S.INV':
      return argc === 1
    case 'LOGNORM.DIST':
      return argc === 3 || argc === 4
    case 'LOGNORM.INV':
      return argc === 3
    case 'MODE':
    case 'MODE.SNGL':
    case 'STDEV':
    case 'STDEV.P':
    case 'STDEV.S':
    case 'STDEVA':
    case 'STDEVP':
    case 'STDEVPA':
    case 'VAR':
    case 'VAR.P':
    case 'VAR.S':
    case 'VARA':
    case 'VARP':
    case 'VARPA':
    case 'SKEW':
    case 'SKEW.P':
    case 'KURT':
    case 'NPV':
      return argc >= 1
    case 'FV':
    case 'PV':
    case 'PMT':
    case 'NPER':
      return argc >= 3 && argc <= 5
    case 'RATE':
      return argc >= 3 && argc <= 6
    case 'IPMT':
    case 'PPMT':
      return argc >= 4 && argc <= 6
    case 'ISPMT':
      return argc === 4
    case 'CUMIPMT':
    case 'CUMPRINC':
      return argc === 6
    case 'DATE':
    case 'TIME':
    case 'DATEDIF':
      return argc === 3
    case 'FVSCHEDULE':
      return argc >= 2
    case 'SLN':
      return argc === 3
    case 'DB':
    case 'DDB':
      return argc === 4 || argc === 5
    case 'SYD':
      return argc === 4
    case 'VDB':
      return argc >= 5 && argc <= 7
    case 'EDATE':
    case 'EOMONTH':
      return argc === 2
    case 'AND':
    case 'OR':
    case 'XOR':
      return argc >= 1
    case 'SWITCH':
      return argc >= 3
    case 'SEQUENCE':
      return argc >= 1 && argc <= 4
    case 'EXPAND':
      return argc >= 2 && argc <= 4
    case 'FILTER':
      return argc === 2 || argc === 3
    case 'UNIQUE':
      return argc >= 1 && argc <= 3
    case 'TRIMRANGE':
      return argc >= 1 && argc <= 3
    case 'OFFSET':
      return argc >= 3 && argc <= 5
    case 'TAKE':
    case 'DROP':
      return argc >= 1 && argc <= 3
    case 'CHOOSECOLS':
    case 'CHOOSEROWS':
      return argc >= 2
    case 'SORT':
      return argc >= 1 && argc <= 4
    case 'SORTBY':
      return argc >= 2
    case 'TOCOL':
    case 'TOROW':
      return argc >= 1 && argc <= 3
    case 'WRAPROWS':
    case 'WRAPCOLS':
      return argc >= 2 && argc <= 4
    case 'LOOKUP':
      return argc === 2 || argc === 3
    case 'AREAS':
    case 'COLUMNS':
    case 'ROWS':
    case 'TRANSPOSE':
      return argc === 1
    case 'HSTACK':
    case 'VSTACK':
      return argc >= 1
    case 'ARRAYTOTEXT':
      return argc === 1 || argc === 2
    case 'MINIFS':
    case 'MAXIFS':
      return argc >= 3 && argc % 2 === 1
    default:
      return true
  }
}

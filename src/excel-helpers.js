// NOTE, THIS FUNCTION IS 1-BASED INDEXING (i.e., COLUMN A IS 1)
export function convertNumToColumnLetter (num) {
  if (num === 0) {
    return ''
  }

  let currentDigit = Math.floor(num / 27)

  if (currentDigit > 0) {
    let remainder = num - currentDigit * 26
    return String.fromCharCode(65 + currentDigit - 1) + convertNumToColumnLetter(remainder)
  } else {
    return String.fromCharCode(65 + num - 1)
  }
}

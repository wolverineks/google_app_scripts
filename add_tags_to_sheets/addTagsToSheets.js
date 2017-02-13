function onOpen () {
  var tags = [
    'Board',
    'Coaching',
    'Communication',
    'Contests',
    'Demographics',
    'Fundraising',
    'Marketing',
    'Merchandise',
    'Members',
    'Music',
    'Performances',
    'Quartets',
    'Recruitment',
    'Rehearsals',
    'Revenue',
    'Sales',
    'Technology',
    'Tickets',
    'Volunteers',
    'Website'
  ]

  var menu = SpreadsheetApp
    .getUi()
    .createMenu('Tags')
  addItemsToMenu(menu, tags)
  menu.addToUi()
}

function addItemsToMenu(menu, tags) {
  for (var i = 0; i < tags.length; i += 1) {
    menu.addItem(tags[i], tags[i].toLowerCase())
  }
}

function board () {
  var tagName = 'board'
  var cell = getActiveCell()
  toggleTag(cell, tagName)
}
function coaching () {
  var tagName = 'coaching'
  var cell = getActiveCell()
  toggleTag(cell, tagName)
}
function communication () {
  var tagName = 'communication'
  var cell = getActiveCell()
  toggleTag(cell, tagName)
}
function contests () {
  var tagName = 'contests'
  var cell = getActiveCell()
  toggleTag(cell, tagName)
}
function demographics () {
  var tagName = 'demographics'
  var cell = getActiveCell()
  toggleTag(cell, tagName)
}
function fundraising () {
  var tagName = 'fundraising'
  var cell = getActiveCell()
  toggleTag(cell, tagName)
}
function marketing () {
  var tagName = 'marketing'
  var cell = getActiveCell()
  toggleTag(cell, tagName)
}
function members () {
  var tagName = 'members'
  var cell = getActiveCell()
  toggleTag(cell, tagName)
}
function merchandise () {
  var tagName = 'merchandise'
  var cell = getActiveCell()
  toggleTag(cell, tagName)
}
function music () {
  var tagName = 'music'
  var cell = getActiveCell()
  toggleTag(cell, tagName)
}
function performances () {
  var tagName = 'performances'
  var cell = getActiveCell()
  toggleTag(cell, tagName)
}
function quartets () {
  var tagName = 'quartets'
  var cell = getActiveCell()
  toggleTag(cell, tagName)
}
function recruitment () {
  var tagName = 'recruitment'
  var cell = getActiveCell()
  toggleTag(cell, tagName)
}
function rehearsals () {
  var tagName = 'rehearsals'
  var cell = getActiveCell()
  toggleTag(cell, tagName)
}
function revenue () {
  var tagName = 'revenue'
  var cell = getActiveCell()
  toggleTag(cell, tagName)
}
function sales () {
  var tagName = 'sales'
  var cell = getActiveCell()
  toggleTag(cell, tagName)
}
function technology () {
  var tagName = 'technology'
  var cell = getActiveCell()
  toggleTag(cell, tagName)
}
function tickets () {
  var tagName = 'tickets'
  var cell = getActiveCell()
  toggleTag(cell, tagName)
}
function volunteers () {
  var tagName = 'volunteers'
  var cell = getActiveCell()
  toggleTag(cell, tagName)
}
function website () {
  var tagName = 'website'
  var cell = getActiveCell()
  toggleTag(cell, tagName)
}

function getActiveCell () {
  var spreadsheet = SpreadsheetApp
    .getActiveSheet()
    .getActiveCell()

  return spreadsheet
}

function toggleTag (cell, tagName) {
  var cellValue = cell.getValue()
  var newValue = ''

  if (includes(cellValue, tagName)) {
    newValue = removeTag(cellValue, tagName)
  } else {
    newValue = addTag(cellValue, tagName)
  }

  cell.setValue(newValue)
}

function includes (cellValue, tagName) {
  return cellValue.indexOf(tagName) >= 0
}

function addTag (cellValue, tagName) {
  var newValue = ''
  if (cellValue == '') {
    newValue = tagName
  } else {
    newValue = cellValue + ', ' + tagName
  }

  return newValue
}

function removeTag (cellValue, tagName) {
  var newValue = ''

  if (isLeading(cellValue, tagName)) {
    newValue = cellValue.replace(tagName + ', ', '')
  } else if (isTrailing(cellValue, tagName)) {
    newValue = cellValue.replace(', ' + tagName, '')
  } else {
    newValue = ''
  }

  return newValue
}

function isLeading (cellValue, tagName) {
  return cellValue.indexOf(tagName + ', ') >= 0
}
function isTrailing (cellValue, tagName) {
  return cellValue.indexOf(', ' + tagName) >= 0
}

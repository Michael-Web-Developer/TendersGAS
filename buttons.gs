function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('Кнопки')
      .addItem('Запустить обмен статусами', 'runTransferStatuses')
      .addItem('Текущий день', 'activeCurrentDay')
      .addItem('Показатели работы', 'setAllQuality')
      .addItem('Обновить информацию строк', 'edit_data')
      .addItem('Закрепить заявку, как избранное', 'favourite_row')
      .addItem('Убрать заявку из избранного', 'remove_favourite_row')
      .addItem('Расчет рек. цены sidebar', 'openSideBar')
      .addItem('Расчет рек. цены model dialog', 'modelDialog')
    .addToUi();
}

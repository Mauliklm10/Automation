---------------------------------------------------------------------------
KeyError                                  Traceback (most recent call last)
~\Anaconda3\lib\site-packages\pandas\core\indexes\base.py in get_loc(self, key, method, tolerance)
   3360             try:
-> 3361                 return self._engine.get_loc(casted_key)
   3362             except KeyError as err:

~\Anaconda3\lib\site-packages\pandas\_libs\index.pyx in pandas._libs.index.IndexEngine.get_loc()

~\Anaconda3\lib\site-packages\pandas\_libs\index.pyx in pandas._libs.index.IndexEngine.get_loc()

pandas\_libs\hashtable_class_helper.pxi in pandas._libs.hashtable.Int64HashTable.get_item()

pandas\_libs\hashtable_class_helper.pxi in pandas._libs.hashtable.Int64HashTable.get_item()

KeyError: 1

The above exception was the direct cause of the following exception:

KeyError                                  Traceback (most recent call last)
~\Anaconda3\lib\site-packages\ipykernel_launcher.py in <module>
     25         start_col = 1
     26 
---> 27     if check_previous_data(existing_data, start_row, start_col, AD.values):
     28         print("Data already present in previous columns for DataFrame 1")
     29     else:

~\Anaconda3\lib\site-packages\ipykernel_launcher.py in check_previous_data(sheet, row, col, data)
      1 def check_previous_data(sheet, row, col, data):
      2     for i in range(len(data)):
----> 3         if sheet.loc[row + i, col] is not None:
      4             return True
      5     return False

~\Anaconda3\lib\site-packages\pandas\core\indexing.py in __getitem__(self, key)
    923                 with suppress(KeyError, IndexError):
    924                     return self.obj._get_value(*key, takeable=self._takeable)
--> 925             return self._getitem_tuple(key)
    926         else:
    927             # we by definition only have the 0th axis

~\Anaconda3\lib\site-packages\pandas\core\indexing.py in _getitem_tuple(self, tup)
   1098     def _getitem_tuple(self, tup: tuple):
   1099         with suppress(IndexingError):
-> 1100             return self._getitem_lowerdim(tup)
   1101 
   1102         # no multi-index, so validate all of the indexers

~\Anaconda3\lib\site-packages\pandas\core\indexing.py in _getitem_lowerdim(self, tup)
    860                     return section
    861                 # This is an elided recursive call to iloc/loc
--> 862                 return getattr(section, self.name)[new_key]
    863 
    864         raise IndexingError("not applicable")

~\Anaconda3\lib\site-packages\pandas\core\indexing.py in __getitem__(self, key)
    929 
    930             maybe_callable = com.apply_if_callable(key, self.obj)
--> 931             return self._getitem_axis(maybe_callable, axis=axis)
    932 
    933     def _is_scalar_access(self, key: tuple):

~\Anaconda3\lib\site-packages\pandas\core\indexing.py in _getitem_axis(self, key, axis)
   1162         # fall thru to straight lookup
   1163         self._validate_key(key, axis)
-> 1164         return self._get_label(key, axis=axis)
   1165 
   1166     def _get_slice_axis(self, slice_obj: slice, axis: int):

~\Anaconda3\lib\site-packages\pandas\core\indexing.py in _get_label(self, label, axis)
   1111     def _get_label(self, label, axis: int):
   1112         # GH#5667 this will fail if the label is not present in the axis.
-> 1113         return self.obj.xs(label, axis=axis)
   1114 
   1115     def _handle_lowerdim_multi_index_axis0(self, tup: tuple):

~\Anaconda3\lib\site-packages\pandas\core\generic.py in xs(self, key, axis, level, drop_level)
   3774                 raise TypeError(f"Expected label or tuple of labels, got {key}") from e
   3775         else:
-> 3776             loc = index.get_loc(key)
   3777 
   3778             if isinstance(loc, np.ndarray):

~\Anaconda3\lib\site-packages\pandas\core\indexes\base.py in get_loc(self, key, method, tolerance)
   3361                 return self._engine.get_loc(casted_key)
   3362             except KeyError as err:
-> 3363                 raise KeyError(key) from err
   3364 
   3365         if is_scalar(key) and isna(key) and not self.hasnans:

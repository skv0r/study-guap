package com.example.autotasks.database;

import android.database.Cursor;
import android.os.CancellationSignal;
import androidx.annotation.NonNull;
import androidx.annotation.Nullable;
import androidx.room.CoroutinesRoom;
import androidx.room.EntityDeletionOrUpdateAdapter;
import androidx.room.EntityInsertionAdapter;
import androidx.room.RoomDatabase;
import androidx.room.RoomSQLiteQuery;
import androidx.room.SharedSQLiteStatement;
import androidx.room.util.CursorUtil;
import androidx.room.util.DBUtil;
import androidx.sqlite.db.SupportSQLiteStatement;
import java.lang.Class;
import java.lang.Exception;
import java.lang.Integer;
import java.lang.Long;
import java.lang.Object;
import java.lang.Override;
import java.lang.String;
import java.lang.SuppressWarnings;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.concurrent.Callable;
import javax.annotation.processing.Generated;
import kotlin.Unit;
import kotlin.coroutines.Continuation;

@Generated("androidx.room.RoomProcessor")
@SuppressWarnings({"unchecked", "deprecation"})
public final class DriverDao_Impl implements DriverDao {
  private final RoomDatabase __db;

  private final EntityInsertionAdapter<Driver> __insertionAdapterOfDriver;

  private final EntityDeletionOrUpdateAdapter<Driver> __deletionAdapterOfDriver;

  private final EntityDeletionOrUpdateAdapter<Driver> __updateAdapterOfDriver;

  private final SharedSQLiteStatement __preparedStmtOfDeleteAllDrivers;

  public DriverDao_Impl(@NonNull final RoomDatabase __db) {
    this.__db = __db;
    this.__insertionAdapterOfDriver = new EntityInsertionAdapter<Driver>(__db) {
      @Override
      @NonNull
      protected String createQuery() {
        return "INSERT OR REPLACE INTO `drivers` (`id`,`full_name`,`driver_number`,`first_name`,`last_name`,`team_name`,`team_colour`,`name_acronym`,`country_code`,`broadcast_name`) VALUES (nullif(?, 0),?,?,?,?,?,?,?,?,?)";
      }

      @Override
      protected void bind(@NonNull final SupportSQLiteStatement statement,
          @NonNull final Driver entity) {
        statement.bindLong(1, entity.getId());
        if (entity.getFullName() == null) {
          statement.bindNull(2);
        } else {
          statement.bindString(2, entity.getFullName());
        }
        statement.bindLong(3, entity.getDriverNumber());
        if (entity.getFirstName() == null) {
          statement.bindNull(4);
        } else {
          statement.bindString(4, entity.getFirstName());
        }
        if (entity.getLastName() == null) {
          statement.bindNull(5);
        } else {
          statement.bindString(5, entity.getLastName());
        }
        if (entity.getTeamName() == null) {
          statement.bindNull(6);
        } else {
          statement.bindString(6, entity.getTeamName());
        }
        if (entity.getTeamColour() == null) {
          statement.bindNull(7);
        } else {
          statement.bindString(7, entity.getTeamColour());
        }
        if (entity.getNameAcronym() == null) {
          statement.bindNull(8);
        } else {
          statement.bindString(8, entity.getNameAcronym());
        }
        if (entity.getCountryCode() == null) {
          statement.bindNull(9);
        } else {
          statement.bindString(9, entity.getCountryCode());
        }
        if (entity.getBroadcastName() == null) {
          statement.bindNull(10);
        } else {
          statement.bindString(10, entity.getBroadcastName());
        }
      }
    };
    this.__deletionAdapterOfDriver = new EntityDeletionOrUpdateAdapter<Driver>(__db) {
      @Override
      @NonNull
      protected String createQuery() {
        return "DELETE FROM `drivers` WHERE `id` = ?";
      }

      @Override
      protected void bind(@NonNull final SupportSQLiteStatement statement,
          @NonNull final Driver entity) {
        statement.bindLong(1, entity.getId());
      }
    };
    this.__updateAdapterOfDriver = new EntityDeletionOrUpdateAdapter<Driver>(__db) {
      @Override
      @NonNull
      protected String createQuery() {
        return "UPDATE OR ABORT `drivers` SET `id` = ?,`full_name` = ?,`driver_number` = ?,`first_name` = ?,`last_name` = ?,`team_name` = ?,`team_colour` = ?,`name_acronym` = ?,`country_code` = ?,`broadcast_name` = ? WHERE `id` = ?";
      }

      @Override
      protected void bind(@NonNull final SupportSQLiteStatement statement,
          @NonNull final Driver entity) {
        statement.bindLong(1, entity.getId());
        if (entity.getFullName() == null) {
          statement.bindNull(2);
        } else {
          statement.bindString(2, entity.getFullName());
        }
        statement.bindLong(3, entity.getDriverNumber());
        if (entity.getFirstName() == null) {
          statement.bindNull(4);
        } else {
          statement.bindString(4, entity.getFirstName());
        }
        if (entity.getLastName() == null) {
          statement.bindNull(5);
        } else {
          statement.bindString(5, entity.getLastName());
        }
        if (entity.getTeamName() == null) {
          statement.bindNull(6);
        } else {
          statement.bindString(6, entity.getTeamName());
        }
        if (entity.getTeamColour() == null) {
          statement.bindNull(7);
        } else {
          statement.bindString(7, entity.getTeamColour());
        }
        if (entity.getNameAcronym() == null) {
          statement.bindNull(8);
        } else {
          statement.bindString(8, entity.getNameAcronym());
        }
        if (entity.getCountryCode() == null) {
          statement.bindNull(9);
        } else {
          statement.bindString(9, entity.getCountryCode());
        }
        if (entity.getBroadcastName() == null) {
          statement.bindNull(10);
        } else {
          statement.bindString(10, entity.getBroadcastName());
        }
        statement.bindLong(11, entity.getId());
      }
    };
    this.__preparedStmtOfDeleteAllDrivers = new SharedSQLiteStatement(__db) {
      @Override
      @NonNull
      public String createQuery() {
        final String _query = "DELETE FROM drivers";
        return _query;
      }
    };
  }

  @Override
  public Object insertDriver(final Driver driver, final Continuation<? super Long> $completion) {
    return CoroutinesRoom.execute(__db, true, new Callable<Long>() {
      @Override
      @NonNull
      public Long call() throws Exception {
        __db.beginTransaction();
        try {
          final Long _result = __insertionAdapterOfDriver.insertAndReturnId(driver);
          __db.setTransactionSuccessful();
          return _result;
        } finally {
          __db.endTransaction();
        }
      }
    }, $completion);
  }

  @Override
  public Object deleteDriver(final Driver driver, final Continuation<? super Unit> $completion) {
    return CoroutinesRoom.execute(__db, true, new Callable<Unit>() {
      @Override
      @NonNull
      public Unit call() throws Exception {
        __db.beginTransaction();
        try {
          __deletionAdapterOfDriver.handle(driver);
          __db.setTransactionSuccessful();
          return Unit.INSTANCE;
        } finally {
          __db.endTransaction();
        }
      }
    }, $completion);
  }

  @Override
  public Object updateDriver(final Driver driver, final Continuation<? super Unit> $completion) {
    return CoroutinesRoom.execute(__db, true, new Callable<Unit>() {
      @Override
      @NonNull
      public Unit call() throws Exception {
        __db.beginTransaction();
        try {
          __updateAdapterOfDriver.handle(driver);
          __db.setTransactionSuccessful();
          return Unit.INSTANCE;
        } finally {
          __db.endTransaction();
        }
      }
    }, $completion);
  }

  @Override
  public Object deleteAllDrivers(final Continuation<? super Unit> $completion) {
    return CoroutinesRoom.execute(__db, true, new Callable<Unit>() {
      @Override
      @NonNull
      public Unit call() throws Exception {
        final SupportSQLiteStatement _stmt = __preparedStmtOfDeleteAllDrivers.acquire();
        try {
          __db.beginTransaction();
          try {
            _stmt.executeUpdateDelete();
            __db.setTransactionSuccessful();
            return Unit.INSTANCE;
          } finally {
            __db.endTransaction();
          }
        } finally {
          __preparedStmtOfDeleteAllDrivers.release(_stmt);
        }
      }
    }, $completion);
  }

  @Override
  public Object getAllDrivers(final Continuation<? super List<Driver>> $completion) {
    final String _sql = "SELECT * FROM drivers ORDER BY driver_number ASC";
    final RoomSQLiteQuery _statement = RoomSQLiteQuery.acquire(_sql, 0);
    final CancellationSignal _cancellationSignal = DBUtil.createCancellationSignal();
    return CoroutinesRoom.execute(__db, false, _cancellationSignal, new Callable<List<Driver>>() {
      @Override
      @NonNull
      public List<Driver> call() throws Exception {
        final Cursor _cursor = DBUtil.query(__db, _statement, false, null);
        try {
          final int _cursorIndexOfId = CursorUtil.getColumnIndexOrThrow(_cursor, "id");
          final int _cursorIndexOfFullName = CursorUtil.getColumnIndexOrThrow(_cursor, "full_name");
          final int _cursorIndexOfDriverNumber = CursorUtil.getColumnIndexOrThrow(_cursor, "driver_number");
          final int _cursorIndexOfFirstName = CursorUtil.getColumnIndexOrThrow(_cursor, "first_name");
          final int _cursorIndexOfLastName = CursorUtil.getColumnIndexOrThrow(_cursor, "last_name");
          final int _cursorIndexOfTeamName = CursorUtil.getColumnIndexOrThrow(_cursor, "team_name");
          final int _cursorIndexOfTeamColour = CursorUtil.getColumnIndexOrThrow(_cursor, "team_colour");
          final int _cursorIndexOfNameAcronym = CursorUtil.getColumnIndexOrThrow(_cursor, "name_acronym");
          final int _cursorIndexOfCountryCode = CursorUtil.getColumnIndexOrThrow(_cursor, "country_code");
          final int _cursorIndexOfBroadcastName = CursorUtil.getColumnIndexOrThrow(_cursor, "broadcast_name");
          final List<Driver> _result = new ArrayList<Driver>(_cursor.getCount());
          while (_cursor.moveToNext()) {
            final Driver _item;
            final int _tmpId;
            _tmpId = _cursor.getInt(_cursorIndexOfId);
            final String _tmpFullName;
            if (_cursor.isNull(_cursorIndexOfFullName)) {
              _tmpFullName = null;
            } else {
              _tmpFullName = _cursor.getString(_cursorIndexOfFullName);
            }
            final int _tmpDriverNumber;
            _tmpDriverNumber = _cursor.getInt(_cursorIndexOfDriverNumber);
            final String _tmpFirstName;
            if (_cursor.isNull(_cursorIndexOfFirstName)) {
              _tmpFirstName = null;
            } else {
              _tmpFirstName = _cursor.getString(_cursorIndexOfFirstName);
            }
            final String _tmpLastName;
            if (_cursor.isNull(_cursorIndexOfLastName)) {
              _tmpLastName = null;
            } else {
              _tmpLastName = _cursor.getString(_cursorIndexOfLastName);
            }
            final String _tmpTeamName;
            if (_cursor.isNull(_cursorIndexOfTeamName)) {
              _tmpTeamName = null;
            } else {
              _tmpTeamName = _cursor.getString(_cursorIndexOfTeamName);
            }
            final String _tmpTeamColour;
            if (_cursor.isNull(_cursorIndexOfTeamColour)) {
              _tmpTeamColour = null;
            } else {
              _tmpTeamColour = _cursor.getString(_cursorIndexOfTeamColour);
            }
            final String _tmpNameAcronym;
            if (_cursor.isNull(_cursorIndexOfNameAcronym)) {
              _tmpNameAcronym = null;
            } else {
              _tmpNameAcronym = _cursor.getString(_cursorIndexOfNameAcronym);
            }
            final String _tmpCountryCode;
            if (_cursor.isNull(_cursorIndexOfCountryCode)) {
              _tmpCountryCode = null;
            } else {
              _tmpCountryCode = _cursor.getString(_cursorIndexOfCountryCode);
            }
            final String _tmpBroadcastName;
            if (_cursor.isNull(_cursorIndexOfBroadcastName)) {
              _tmpBroadcastName = null;
            } else {
              _tmpBroadcastName = _cursor.getString(_cursorIndexOfBroadcastName);
            }
            _item = new Driver(_tmpId,_tmpFullName,_tmpDriverNumber,_tmpFirstName,_tmpLastName,_tmpTeamName,_tmpTeamColour,_tmpNameAcronym,_tmpCountryCode,_tmpBroadcastName);
            _result.add(_item);
          }
          return _result;
        } finally {
          _cursor.close();
          _statement.release();
        }
      }
    }, $completion);
  }

  @Override
  public Object getDriversCount(final Continuation<? super Integer> $completion) {
    final String _sql = "SELECT COUNT(*) FROM drivers";
    final RoomSQLiteQuery _statement = RoomSQLiteQuery.acquire(_sql, 0);
    final CancellationSignal _cancellationSignal = DBUtil.createCancellationSignal();
    return CoroutinesRoom.execute(__db, false, _cancellationSignal, new Callable<Integer>() {
      @Override
      @NonNull
      public Integer call() throws Exception {
        final Cursor _cursor = DBUtil.query(__db, _statement, false, null);
        try {
          final Integer _result;
          if (_cursor.moveToFirst()) {
            final Integer _tmp;
            if (_cursor.isNull(0)) {
              _tmp = null;
            } else {
              _tmp = _cursor.getInt(0);
            }
            _result = _tmp;
          } else {
            _result = null;
          }
          return _result;
        } finally {
          _cursor.close();
          _statement.release();
        }
      }
    }, $completion);
  }

  @Override
  public Object getDriverById(final int id, final Continuation<? super Driver> $completion) {
    final String _sql = "SELECT * FROM drivers WHERE id = ?";
    final RoomSQLiteQuery _statement = RoomSQLiteQuery.acquire(_sql, 1);
    int _argIndex = 1;
    _statement.bindLong(_argIndex, id);
    final CancellationSignal _cancellationSignal = DBUtil.createCancellationSignal();
    return CoroutinesRoom.execute(__db, false, _cancellationSignal, new Callable<Driver>() {
      @Override
      @Nullable
      public Driver call() throws Exception {
        final Cursor _cursor = DBUtil.query(__db, _statement, false, null);
        try {
          final int _cursorIndexOfId = CursorUtil.getColumnIndexOrThrow(_cursor, "id");
          final int _cursorIndexOfFullName = CursorUtil.getColumnIndexOrThrow(_cursor, "full_name");
          final int _cursorIndexOfDriverNumber = CursorUtil.getColumnIndexOrThrow(_cursor, "driver_number");
          final int _cursorIndexOfFirstName = CursorUtil.getColumnIndexOrThrow(_cursor, "first_name");
          final int _cursorIndexOfLastName = CursorUtil.getColumnIndexOrThrow(_cursor, "last_name");
          final int _cursorIndexOfTeamName = CursorUtil.getColumnIndexOrThrow(_cursor, "team_name");
          final int _cursorIndexOfTeamColour = CursorUtil.getColumnIndexOrThrow(_cursor, "team_colour");
          final int _cursorIndexOfNameAcronym = CursorUtil.getColumnIndexOrThrow(_cursor, "name_acronym");
          final int _cursorIndexOfCountryCode = CursorUtil.getColumnIndexOrThrow(_cursor, "country_code");
          final int _cursorIndexOfBroadcastName = CursorUtil.getColumnIndexOrThrow(_cursor, "broadcast_name");
          final Driver _result;
          if (_cursor.moveToFirst()) {
            final int _tmpId;
            _tmpId = _cursor.getInt(_cursorIndexOfId);
            final String _tmpFullName;
            if (_cursor.isNull(_cursorIndexOfFullName)) {
              _tmpFullName = null;
            } else {
              _tmpFullName = _cursor.getString(_cursorIndexOfFullName);
            }
            final int _tmpDriverNumber;
            _tmpDriverNumber = _cursor.getInt(_cursorIndexOfDriverNumber);
            final String _tmpFirstName;
            if (_cursor.isNull(_cursorIndexOfFirstName)) {
              _tmpFirstName = null;
            } else {
              _tmpFirstName = _cursor.getString(_cursorIndexOfFirstName);
            }
            final String _tmpLastName;
            if (_cursor.isNull(_cursorIndexOfLastName)) {
              _tmpLastName = null;
            } else {
              _tmpLastName = _cursor.getString(_cursorIndexOfLastName);
            }
            final String _tmpTeamName;
            if (_cursor.isNull(_cursorIndexOfTeamName)) {
              _tmpTeamName = null;
            } else {
              _tmpTeamName = _cursor.getString(_cursorIndexOfTeamName);
            }
            final String _tmpTeamColour;
            if (_cursor.isNull(_cursorIndexOfTeamColour)) {
              _tmpTeamColour = null;
            } else {
              _tmpTeamColour = _cursor.getString(_cursorIndexOfTeamColour);
            }
            final String _tmpNameAcronym;
            if (_cursor.isNull(_cursorIndexOfNameAcronym)) {
              _tmpNameAcronym = null;
            } else {
              _tmpNameAcronym = _cursor.getString(_cursorIndexOfNameAcronym);
            }
            final String _tmpCountryCode;
            if (_cursor.isNull(_cursorIndexOfCountryCode)) {
              _tmpCountryCode = null;
            } else {
              _tmpCountryCode = _cursor.getString(_cursorIndexOfCountryCode);
            }
            final String _tmpBroadcastName;
            if (_cursor.isNull(_cursorIndexOfBroadcastName)) {
              _tmpBroadcastName = null;
            } else {
              _tmpBroadcastName = _cursor.getString(_cursorIndexOfBroadcastName);
            }
            _result = new Driver(_tmpId,_tmpFullName,_tmpDriverNumber,_tmpFirstName,_tmpLastName,_tmpTeamName,_tmpTeamColour,_tmpNameAcronym,_tmpCountryCode,_tmpBroadcastName);
          } else {
            _result = null;
          }
          return _result;
        } finally {
          _cursor.close();
          _statement.release();
        }
      }
    }, $completion);
  }

  @Override
  public Object getDriverByNumber(final int driverNumber,
      final Continuation<? super Driver> $completion) {
    final String _sql = "SELECT * FROM drivers WHERE driver_number = ?";
    final RoomSQLiteQuery _statement = RoomSQLiteQuery.acquire(_sql, 1);
    int _argIndex = 1;
    _statement.bindLong(_argIndex, driverNumber);
    final CancellationSignal _cancellationSignal = DBUtil.createCancellationSignal();
    return CoroutinesRoom.execute(__db, false, _cancellationSignal, new Callable<Driver>() {
      @Override
      @Nullable
      public Driver call() throws Exception {
        final Cursor _cursor = DBUtil.query(__db, _statement, false, null);
        try {
          final int _cursorIndexOfId = CursorUtil.getColumnIndexOrThrow(_cursor, "id");
          final int _cursorIndexOfFullName = CursorUtil.getColumnIndexOrThrow(_cursor, "full_name");
          final int _cursorIndexOfDriverNumber = CursorUtil.getColumnIndexOrThrow(_cursor, "driver_number");
          final int _cursorIndexOfFirstName = CursorUtil.getColumnIndexOrThrow(_cursor, "first_name");
          final int _cursorIndexOfLastName = CursorUtil.getColumnIndexOrThrow(_cursor, "last_name");
          final int _cursorIndexOfTeamName = CursorUtil.getColumnIndexOrThrow(_cursor, "team_name");
          final int _cursorIndexOfTeamColour = CursorUtil.getColumnIndexOrThrow(_cursor, "team_colour");
          final int _cursorIndexOfNameAcronym = CursorUtil.getColumnIndexOrThrow(_cursor, "name_acronym");
          final int _cursorIndexOfCountryCode = CursorUtil.getColumnIndexOrThrow(_cursor, "country_code");
          final int _cursorIndexOfBroadcastName = CursorUtil.getColumnIndexOrThrow(_cursor, "broadcast_name");
          final Driver _result;
          if (_cursor.moveToFirst()) {
            final int _tmpId;
            _tmpId = _cursor.getInt(_cursorIndexOfId);
            final String _tmpFullName;
            if (_cursor.isNull(_cursorIndexOfFullName)) {
              _tmpFullName = null;
            } else {
              _tmpFullName = _cursor.getString(_cursorIndexOfFullName);
            }
            final int _tmpDriverNumber;
            _tmpDriverNumber = _cursor.getInt(_cursorIndexOfDriverNumber);
            final String _tmpFirstName;
            if (_cursor.isNull(_cursorIndexOfFirstName)) {
              _tmpFirstName = null;
            } else {
              _tmpFirstName = _cursor.getString(_cursorIndexOfFirstName);
            }
            final String _tmpLastName;
            if (_cursor.isNull(_cursorIndexOfLastName)) {
              _tmpLastName = null;
            } else {
              _tmpLastName = _cursor.getString(_cursorIndexOfLastName);
            }
            final String _tmpTeamName;
            if (_cursor.isNull(_cursorIndexOfTeamName)) {
              _tmpTeamName = null;
            } else {
              _tmpTeamName = _cursor.getString(_cursorIndexOfTeamName);
            }
            final String _tmpTeamColour;
            if (_cursor.isNull(_cursorIndexOfTeamColour)) {
              _tmpTeamColour = null;
            } else {
              _tmpTeamColour = _cursor.getString(_cursorIndexOfTeamColour);
            }
            final String _tmpNameAcronym;
            if (_cursor.isNull(_cursorIndexOfNameAcronym)) {
              _tmpNameAcronym = null;
            } else {
              _tmpNameAcronym = _cursor.getString(_cursorIndexOfNameAcronym);
            }
            final String _tmpCountryCode;
            if (_cursor.isNull(_cursorIndexOfCountryCode)) {
              _tmpCountryCode = null;
            } else {
              _tmpCountryCode = _cursor.getString(_cursorIndexOfCountryCode);
            }
            final String _tmpBroadcastName;
            if (_cursor.isNull(_cursorIndexOfBroadcastName)) {
              _tmpBroadcastName = null;
            } else {
              _tmpBroadcastName = _cursor.getString(_cursorIndexOfBroadcastName);
            }
            _result = new Driver(_tmpId,_tmpFullName,_tmpDriverNumber,_tmpFirstName,_tmpLastName,_tmpTeamName,_tmpTeamColour,_tmpNameAcronym,_tmpCountryCode,_tmpBroadcastName);
          } else {
            _result = null;
          }
          return _result;
        } finally {
          _cursor.close();
          _statement.release();
        }
      }
    }, $completion);
  }

  @NonNull
  public static List<Class<?>> getRequiredConverters() {
    return Collections.emptyList();
  }
}

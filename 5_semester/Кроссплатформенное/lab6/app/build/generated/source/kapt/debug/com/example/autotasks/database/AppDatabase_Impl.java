package com.example.autotasks.database;

import androidx.annotation.NonNull;
import androidx.room.DatabaseConfiguration;
import androidx.room.InvalidationTracker;
import androidx.room.RoomDatabase;
import androidx.room.RoomOpenHelper;
import androidx.room.migration.AutoMigrationSpec;
import androidx.room.migration.Migration;
import androidx.room.util.DBUtil;
import androidx.room.util.TableInfo;
import androidx.sqlite.db.SupportSQLiteDatabase;
import androidx.sqlite.db.SupportSQLiteOpenHelper;
import java.lang.Class;
import java.lang.Override;
import java.lang.String;
import java.lang.SuppressWarnings;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;
import javax.annotation.processing.Generated;

@Generated("androidx.room.RoomProcessor")
@SuppressWarnings({"unchecked", "deprecation"})
public final class AppDatabase_Impl extends AppDatabase {
  private volatile DriverDao _driverDao;

  @Override
  @NonNull
  protected SupportSQLiteOpenHelper createOpenHelper(@NonNull final DatabaseConfiguration config) {
    final SupportSQLiteOpenHelper.Callback _openCallback = new RoomOpenHelper(config, new RoomOpenHelper.Delegate(2) {
      @Override
      public void createAllTables(@NonNull final SupportSQLiteDatabase db) {
        db.execSQL("CREATE TABLE IF NOT EXISTS `drivers` (`id` INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL, `full_name` TEXT NOT NULL, `driver_number` INTEGER NOT NULL, `first_name` TEXT NOT NULL, `last_name` TEXT NOT NULL, `team_name` TEXT NOT NULL, `team_colour` TEXT NOT NULL, `name_acronym` TEXT NOT NULL, `country_code` TEXT NOT NULL, `broadcast_name` TEXT NOT NULL)");
        db.execSQL("CREATE UNIQUE INDEX IF NOT EXISTS `index_drivers_driver_number` ON `drivers` (`driver_number`)");
        db.execSQL("CREATE TABLE IF NOT EXISTS room_master_table (id INTEGER PRIMARY KEY,identity_hash TEXT)");
        db.execSQL("INSERT OR REPLACE INTO room_master_table (id,identity_hash) VALUES(42, '46f8592963c711afddebf1d9edab044c')");
      }

      @Override
      public void dropAllTables(@NonNull final SupportSQLiteDatabase db) {
        db.execSQL("DROP TABLE IF EXISTS `drivers`");
        final List<? extends RoomDatabase.Callback> _callbacks = mCallbacks;
        if (_callbacks != null) {
          for (RoomDatabase.Callback _callback : _callbacks) {
            _callback.onDestructiveMigration(db);
          }
        }
      }

      @Override
      public void onCreate(@NonNull final SupportSQLiteDatabase db) {
        final List<? extends RoomDatabase.Callback> _callbacks = mCallbacks;
        if (_callbacks != null) {
          for (RoomDatabase.Callback _callback : _callbacks) {
            _callback.onCreate(db);
          }
        }
      }

      @Override
      public void onOpen(@NonNull final SupportSQLiteDatabase db) {
        mDatabase = db;
        internalInitInvalidationTracker(db);
        final List<? extends RoomDatabase.Callback> _callbacks = mCallbacks;
        if (_callbacks != null) {
          for (RoomDatabase.Callback _callback : _callbacks) {
            _callback.onOpen(db);
          }
        }
      }

      @Override
      public void onPreMigrate(@NonNull final SupportSQLiteDatabase db) {
        DBUtil.dropFtsSyncTriggers(db);
      }

      @Override
      public void onPostMigrate(@NonNull final SupportSQLiteDatabase db) {
      }

      @Override
      @NonNull
      public RoomOpenHelper.ValidationResult onValidateSchema(
          @NonNull final SupportSQLiteDatabase db) {
        final HashMap<String, TableInfo.Column> _columnsDrivers = new HashMap<String, TableInfo.Column>(10);
        _columnsDrivers.put("id", new TableInfo.Column("id", "INTEGER", true, 1, null, TableInfo.CREATED_FROM_ENTITY));
        _columnsDrivers.put("full_name", new TableInfo.Column("full_name", "TEXT", true, 0, null, TableInfo.CREATED_FROM_ENTITY));
        _columnsDrivers.put("driver_number", new TableInfo.Column("driver_number", "INTEGER", true, 0, null, TableInfo.CREATED_FROM_ENTITY));
        _columnsDrivers.put("first_name", new TableInfo.Column("first_name", "TEXT", true, 0, null, TableInfo.CREATED_FROM_ENTITY));
        _columnsDrivers.put("last_name", new TableInfo.Column("last_name", "TEXT", true, 0, null, TableInfo.CREATED_FROM_ENTITY));
        _columnsDrivers.put("team_name", new TableInfo.Column("team_name", "TEXT", true, 0, null, TableInfo.CREATED_FROM_ENTITY));
        _columnsDrivers.put("team_colour", new TableInfo.Column("team_colour", "TEXT", true, 0, null, TableInfo.CREATED_FROM_ENTITY));
        _columnsDrivers.put("name_acronym", new TableInfo.Column("name_acronym", "TEXT", true, 0, null, TableInfo.CREATED_FROM_ENTITY));
        _columnsDrivers.put("country_code", new TableInfo.Column("country_code", "TEXT", true, 0, null, TableInfo.CREATED_FROM_ENTITY));
        _columnsDrivers.put("broadcast_name", new TableInfo.Column("broadcast_name", "TEXT", true, 0, null, TableInfo.CREATED_FROM_ENTITY));
        final HashSet<TableInfo.ForeignKey> _foreignKeysDrivers = new HashSet<TableInfo.ForeignKey>(0);
        final HashSet<TableInfo.Index> _indicesDrivers = new HashSet<TableInfo.Index>(1);
        _indicesDrivers.add(new TableInfo.Index("index_drivers_driver_number", true, Arrays.asList("driver_number"), Arrays.asList("ASC")));
        final TableInfo _infoDrivers = new TableInfo("drivers", _columnsDrivers, _foreignKeysDrivers, _indicesDrivers);
        final TableInfo _existingDrivers = TableInfo.read(db, "drivers");
        if (!_infoDrivers.equals(_existingDrivers)) {
          return new RoomOpenHelper.ValidationResult(false, "drivers(com.example.autotasks.database.Driver).\n"
                  + " Expected:\n" + _infoDrivers + "\n"
                  + " Found:\n" + _existingDrivers);
        }
        return new RoomOpenHelper.ValidationResult(true, null);
      }
    }, "46f8592963c711afddebf1d9edab044c", "2a77f3c69054595755bb621df786acfa");
    final SupportSQLiteOpenHelper.Configuration _sqliteConfig = SupportSQLiteOpenHelper.Configuration.builder(config.context).name(config.name).callback(_openCallback).build();
    final SupportSQLiteOpenHelper _helper = config.sqliteOpenHelperFactory.create(_sqliteConfig);
    return _helper;
  }

  @Override
  @NonNull
  protected InvalidationTracker createInvalidationTracker() {
    final HashMap<String, String> _shadowTablesMap = new HashMap<String, String>(0);
    final HashMap<String, Set<String>> _viewTables = new HashMap<String, Set<String>>(0);
    return new InvalidationTracker(this, _shadowTablesMap, _viewTables, "drivers");
  }

  @Override
  public void clearAllTables() {
    super.assertNotMainThread();
    final SupportSQLiteDatabase _db = super.getOpenHelper().getWritableDatabase();
    try {
      super.beginTransaction();
      _db.execSQL("DELETE FROM `drivers`");
      super.setTransactionSuccessful();
    } finally {
      super.endTransaction();
      _db.query("PRAGMA wal_checkpoint(FULL)").close();
      if (!_db.inTransaction()) {
        _db.execSQL("VACUUM");
      }
    }
  }

  @Override
  @NonNull
  protected Map<Class<?>, List<Class<?>>> getRequiredTypeConverters() {
    final HashMap<Class<?>, List<Class<?>>> _typeConvertersMap = new HashMap<Class<?>, List<Class<?>>>();
    _typeConvertersMap.put(DriverDao.class, DriverDao_Impl.getRequiredConverters());
    return _typeConvertersMap;
  }

  @Override
  @NonNull
  public Set<Class<? extends AutoMigrationSpec>> getRequiredAutoMigrationSpecs() {
    final HashSet<Class<? extends AutoMigrationSpec>> _autoMigrationSpecsSet = new HashSet<Class<? extends AutoMigrationSpec>>();
    return _autoMigrationSpecsSet;
  }

  @Override
  @NonNull
  public List<Migration> getAutoMigrations(
      @NonNull final Map<Class<? extends AutoMigrationSpec>, AutoMigrationSpec> autoMigrationSpecs) {
    final List<Migration> _autoMigrations = new ArrayList<Migration>();
    return _autoMigrations;
  }

  @Override
  public DriverDao driverDao() {
    if (_driverDao != null) {
      return _driverDao;
    } else {
      synchronized(this) {
        if(_driverDao == null) {
          _driverDao = new DriverDao_Impl(this);
        }
        return _driverDao;
      }
    }
  }
}

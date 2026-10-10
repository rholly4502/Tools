/* ============================================================================
   Update verifikasi WhatsApp dari SAP.wa_sap_10102026

   Sumber : SAP.wa_sap_10102026 (CustomerProfileID NOT NULL, value_fixed = WA no)

   Target :
     1. CustomerCDP.Verification
          WhatsAppNo = value_fixed AND CustomerProfileID = pemilik -> WAVerifiedStatus = 'VerifyIn'
          WhatsAppNo sama, CustomerProfileID lain                 -> WAVerifiedStatus = 'False'
          (hanya diproses kalau baris pemilik ADA di Verification)
     2. CustomerCDP.WhatsAppCDP
          CustomerProfileID = pemilik AND WhatsAppNo = value_fixed -> WAVerifiedStaus = 1
          WhatsAppNo sama, CustomerProfileID lain                  -> WAVerifiedStaus = 0
          (Catatan: WAVerifiedStaus = nama actual kolom, bukan typo)

   Tabel log (permanen, simpan LAMA + BARU + ID, untuk rollback):
     CustomerCDP._TmpWAVerify_Verification_10102026
     CustomerCDP._TmpWAVerify_WhatsAppCDP_10102026

   Mode:
     PREVIEW  = bikin tabel log + ringkasan, TIDAK mengubah data asli
     EXECUTE  = update data asli per batch berdasarkan tabel log
     ROLLBACK = kembalikan data asli ke nilai lama dari tabel log
============================================================================ */
SET NOCOUNT ON;
SET XACT_ABORT ON;

------------------------------------------------------------
-- Parameter
------------------------------------------------------------
DECLARE @Mode        VARCHAR(10) = 'PREVIEW';  -- PREVIEW | EXECUTE | ROLLBACK
DECLARE @RebuildLog  BIT         = 1;          -- 1 = drop & bikin ulang tabel log (ditolak kalau IsUpdated = 1)
DECLARE @BatchSize   INT         = 10000;
DECLARE @By          VARCHAR(50) = 'ScriptWAVerify';
DECLARE @Today       DATETIME    = CAST(SYSDATETIMEOFFSET() AT TIME ZONE 'SE Asia Standard Time' AS DATETIME);
DECLARE @TrueStatus  VARCHAR(20) = 'VerifyIn';
DECLARE @FalseStatus VARCHAR(20) = 'False';

-- Validasi mode
IF @Mode NOT IN ('PREVIEW','EXECUTE','ROLLBACK')
    THROW 50001, 'Mode tidak valid. Gunakan PREVIEW, EXECUTE, atau ROLLBACK.', 1;

DECLARE @From INT, @To INT, @Max INT, @n INT, @Total BIGINT, @Msg VARCHAR(400);

------------------------------------------------------------
-- ROLLBACK: kembalikan nilai lama dari tabel log
------------------------------------------------------------
IF @Mode = 'ROLLBACK'
BEGIN
    IF OBJECT_ID('CustomerCDP._TmpWAVerify_Verification_10102026','U') IS NULL
       OR OBJECT_ID('CustomerCDP._TmpWAVerify_WhatsAppCDP_10102026','U') IS NULL
        THROW 50003, 'Tabel log tidak lengkap, rollback dibatalkan.', 1;

    -- urutan terbalik dari EXECUTE: WhatsAppCDP -> Verification

    ---------------- WhatsAppCDP ----------------
    SELECT @Max = ISNULL(MAX(RowNum), 0) FROM CustomerCDP._TmpWAVerify_WhatsAppCDP_10102026;
    SET @From = 1; SET @Total = 0;
    WHILE @From <= @Max
    BEGIN
        SET @To = @From + @BatchSize - 1;
        BEGIN TRY
            BEGIN TRAN;

            UPDATE w
            SET w.WAVerifiedStaus       = l.OldStatus,
                w.LastModifiedBy        = l.OldLastModifiedBy,
                w.LastModifiedTime      = l.OldLastModifiedTime
            FROM CustomerCDP.WhatsAppCDP w
            JOIN CustomerCDP._TmpWAVerify_WhatsAppCDP_10102026 l ON l.WhatsAppID = w.ID
            WHERE l.RowNum BETWEEN @From AND @To
              AND l.IsUpdated = 1
              -- guard: hanya kalau nilai saat ini masih sama dengan hasil script
              AND ISNULL(CAST(w.WAVerifiedStaus AS INT), -1) = ISNULL(CAST(l.NewStatus AS INT), -1);

            SET @n = @@ROWCOUNT;

            UPDATE CustomerCDP._TmpWAVerify_WhatsAppCDP_10102026
            SET IsUpdated = 0, RolledBackTime = @Today
            WHERE RowNum BETWEEN @From AND @To AND IsUpdated = 1;

            COMMIT TRAN;
        END TRY
        BEGIN CATCH
            IF XACT_STATE() <> 0 ROLLBACK TRAN;
            SET @Msg = CONVERT(VARCHAR(30), GETDATE(), 120) + ' - ERROR rollback WhatsAppCDP RowNum ' + CONVERT(VARCHAR(20), @From) + ': ' + ERROR_MESSAGE();
            RAISERROR(@Msg, 0, 1) WITH NOWAIT;
            THROW;
        END CATCH

        SET @Total += @n;
        SET @Msg = CONVERT(VARCHAR(30), GETDATE(), 120) + ' - rollback WhatsAppCDP RowNum ' + CONVERT(VARCHAR(20), @From)
                 + ' s/d ' + CONVERT(VARCHAR(20), CASE WHEN @To > @Max THEN @Max ELSE @To END) + ' dari ' + CONVERT(VARCHAR(20), @Max)
                 + ' | dikembalikan batch ini: ' + CONVERT(VARCHAR(20), @n) + ' | total: ' + CONVERT(VARCHAR(20), @Total);
        RAISERROR(@Msg, 0, 1) WITH NOWAIT;
        SET @From = @To + 1;
    END

    ---------------- Verification ----------------
    SELECT @Max = ISNULL(MAX(RowNum), 0) FROM CustomerCDP._TmpWAVerify_Verification_10102026;
    SET @From = 1; SET @Total = 0;
    WHILE @From <= @Max
    BEGIN
        SET @To = @From + @BatchSize - 1;
        BEGIN TRY
            BEGIN TRAN;

            UPDATE v
            SET v.WAVerifiedStatus   = l.OldStatus,
                v.LastModifiedBy     = l.OldLastModifiedBy,
                v.LastModifiedTime   = l.OldLastModifiedTime
            FROM CustomerCDP.Verification v
            JOIN CustomerCDP._TmpWAVerify_Verification_10102026 l ON l.VerificationID = v.ID
            WHERE l.RowNum BETWEEN @From AND @To
              AND l.IsUpdated = 1
              AND ISNULL(CAST(v.WAVerifiedStatus AS VARCHAR(20)), '') = ISNULL(l.NewStatus, '');

            SET @n = @@ROWCOUNT;

            UPDATE CustomerCDP._TmpWAVerify_Verification_10102026
            SET IsUpdated = 0, RolledBackTime = @Today
            WHERE RowNum BETWEEN @From AND @To AND IsUpdated = 1;

            COMMIT TRAN;
        END TRY
        BEGIN CATCH
            IF XACT_STATE() <> 0 ROLLBACK TRAN;
            SET @Msg = CONVERT(VARCHAR(30), GETDATE(), 120) + ' - ERROR rollback Verification RowNum ' + CONVERT(VARCHAR(20), @From) + ': ' + ERROR_MESSAGE();
            RAISERROR(@Msg, 0, 1) WITH NOWAIT;
            THROW;
        END CATCH

        SET @Total += @n;
        SET @Msg = CONVERT(VARCHAR(30), GETDATE(), 120) + ' - rollback Verification RowNum ' + CONVERT(VARCHAR(20), @From)
                 + ' s/d ' + CONVERT(VARCHAR(20), CASE WHEN @To > @Max THEN @Max ELSE @To END) + ' dari ' + CONVERT(VARCHAR(20), @Max)
                 + ' | dikembalikan batch ini: ' + CONVERT(VARCHAR(20), @n) + ' | total: ' + CONVERT(VARCHAR(20), @Total);
        RAISERROR(@Msg, 0, 1) WITH NOWAIT;
        SET @From = @To + 1;
    END

    RAISERROR('ROLLBACK SELESAI.', 0, 1) WITH NOWAIT;

    -- cleanup temp
    IF OBJECT_ID('tempdb..#Src')      IS NOT NULL DROP TABLE #Src;
    IF OBJECT_ID('tempdb..#Valid')    IS NOT NULL DROP TABLE #Valid;
    IF OBJECT_ID('tempdb..#VerOwned') IS NOT NULL DROP TABLE #VerOwned;
    IF OBJECT_ID('tempdb..#WaOwned')  IS NOT NULL DROP TABLE #WaOwned;

    RETURN;  -- CRITICAL: jangan lanjut ke PREVIEW/EXECUTE
END

------------------------------------------------------------
-- PREVIEW: bikin tabel log + ringkasan, TIDAK ubah data asli
------------------------------------------------------------
IF @Mode = 'PREVIEW'
BEGIN
    -- 1. Build source temp tables (dedup konflik)
    IF OBJECT_ID('tempdb..#Src')      IS NOT NULL DROP TABLE #Src;
    IF OBJECT_ID('tempdb..#Valid')    IS NOT NULL DROP TABLE #Valid;
    IF OBJECT_ID('tempdb..#VerOwned') IS NOT NULL DROP TABLE #VerOwned;
    IF OBJECT_ID('tempdb..#WaOwned')  IS NOT NULL DROP TABLE #WaOwned;

    -- Source: filter CustomerProfileID NOT NULL, value_fixed tidak kosong
    SELECT CustomerProfileID, LTRIM(RTRIM(value_fixed)) AS WhatsAppNo
    INTO #Src
    FROM SAP.wa_sap_10102026
    WHERE CustomerProfileID IS NOT NULL
      AND value_fixed IS NOT NULL
      AND LTRIM(RTRIM(value_fixed)) <> '';

    -- Valid: exclude konflik (1 nomor -> >1 profile OR 1 profile -> >1 nomor)
    SELECT s.CustomerProfileID, s.WhatsAppNo
    INTO #Valid
    FROM #Src s
    WHERE NOT EXISTS (
        SELECT 1 FROM #Src s2
        WHERE s2.WhatsAppNo = s.WhatsAppNo
          AND s2.CustomerProfileID <> s.CustomerProfileID
    )
    AND NOT EXISTS (
        SELECT 1 FROM #Src s3
        WHERE s3.CustomerProfileID = s.CustomerProfileID
          AND s3.WhatsAppNo <> s.WhatsAppNo
    );

    DECLARE @TotalSrc INT, @ValidCnt INT, @ConflictCnt INT;
    SELECT @TotalSrc = COUNT(*) FROM #Src;
    SELECT @ValidCnt = COUNT(*) FROM #Valid;
    SET @ConflictCnt = @TotalSrc - @ValidCnt;

    RAISERROR('Source: %d baris | Valid: %d | Konflik (dilewati): %d', 0, 1, @TotalSrc, @ValidCnt, @ConflictCnt) WITH NOWAIT;

    -- 2. Create / rebuild log tables
    IF @RebuildLog = 1
    BEGIN
        IF OBJECT_ID('CustomerCDP._TmpWAVerify_Verification_10102026','U') IS NOT NULL
           AND EXISTS (SELECT 1 FROM CustomerCDP._TmpWAVerify_Verification_10102026 WHERE IsUpdated = 1)
            THROW 50002, 'Tidak bisa rebuild log Verification: masih ada IsUpdated=1. ROLLBACK dulu atau set @RebuildLog=0.', 1;

        IF OBJECT_ID('CustomerCDP._TmpWAVerify_WhatsAppCDP_10102026','U') IS NOT NULL
           AND EXISTS (SELECT 1 FROM CustomerCDP._TmpWAVerify_WhatsAppCDP_10102026 WHERE IsUpdated = 1)
            THROW 50002, 'Tidak bisa rebuild log WhatsAppCDP: masih ada IsUpdated=1. ROLLBACK dulu atau set @RebuildLog=0.', 1;

        IF OBJECT_ID('CustomerCDP._TmpWAVerify_Verification_10102026','U') IS NOT NULL
            DROP TABLE CustomerCDP._TmpWAVerify_Verification_10102026;
        IF OBJECT_ID('CustomerCDP._TmpWAVerify_WhatsAppCDP_10102026','U') IS NOT NULL
            DROP TABLE CustomerCDP._TmpWAVerify_WhatsAppCDP_10102026;
    END

    -- Log: Verification
    IF OBJECT_ID('CustomerCDP._TmpWAVerify_Verification_10102026','U') IS NULL
    BEGIN
        CREATE TABLE CustomerCDP._TmpWAVerify_Verification_10102026 (
            RowNum              INT IDENTITY(1,1) PRIMARY KEY,
            VerificationID      BIGINT NOT NULL,
            CustomerProfileID   BIGINT NULL,
            WhatsAppNo          VARCHAR(50) NULL,
            OldStatus           VARCHAR(20) NULL,
            NewStatus           VARCHAR(20) NULL,
            OldLastModifiedBy   VARCHAR(100) NULL,
            OldLastModifiedTime DATETIME NULL,
            NeedUpdate          BIT NOT NULL DEFAULT 0,
            IsUpdated           BIT NOT NULL DEFAULT 0,
            UpdatedTime         DATETIME NULL,
            RolledBackTime      DATETIME NULL
        );
    END
    ELSE
        TRUNCATE TABLE CustomerCDP._TmpWAVerify_Verification_10102026;

    -- Log: WhatsAppCDP (WAVerifiedStaus = nama actual kolom, bukan typo)
    IF OBJECT_ID('CustomerCDP._TmpWAVerify_WhatsAppCDP_10102026','U') IS NULL
    BEGIN
        CREATE TABLE CustomerCDP._TmpWAVerify_WhatsAppCDP_10102026 (
            RowNum              INT IDENTITY(1,1) PRIMARY KEY,
            WhatsAppID          BIGINT NOT NULL,
            CustomerProfileID   BIGINT NULL,
            WhatsAppNo          VARCHAR(50) NULL,
            OldStatus           INT NULL,
            NewStatus           INT NULL,
            OldLastModifiedBy   VARCHAR(100) NULL,
            OldLastModifiedTime DATETIME NULL,
            NeedUpdate          BIT NOT NULL DEFAULT 0,
            IsUpdated           BIT NOT NULL DEFAULT 0,
            UpdatedTime         DATETIME NULL,
            RolledBackTime      DATETIME NULL
        );
    END
    ELSE
        TRUNCATE TABLE CustomerCDP._TmpWAVerify_WhatsAppCDP_10102026;

    -- 3. Populate Verification log — OWNER rows (VerifyIn)
    INSERT INTO CustomerCDP._TmpWAVerify_Verification_10102026
        (VerificationID, CustomerProfileID, WhatsAppNo, OldStatus, NewStatus, OldLastModifiedBy, OldLastModifiedTime, NeedUpdate)
    SELECT v.ID, v.CustomerProfileID, v.WhatsAppNo,
           ISNULL(CAST(v.WAVerifiedStatus AS VARCHAR(20)), ''),
           @TrueStatus,
           v.LastModifiedBy, v.LastModifiedTime,
           CASE WHEN ISNULL(CAST(v.WAVerifiedStatus AS VARCHAR(20)), '') <> @TrueStatus THEN 1 ELSE 0 END
    FROM CustomerCDP.Verification v
    JOIN #Valid s ON v.CustomerProfileID = s.CustomerProfileID AND v.WhatsAppNo = s.WhatsAppNo;

    -- #VerOwned: nomor yang punya baris pemilik di Verification
    SELECT DISTINCT s.WhatsAppNo
    INTO #VerOwned
    FROM #Valid s
    WHERE EXISTS (
        SELECT 1 FROM CustomerCDP.Verification v
        WHERE v.CustomerProfileID = s.CustomerProfileID AND v.WhatsAppNo = s.WhatsAppNo
    );

    -- Verification log — OTHER rows (False), hanya kalau owner ada
    INSERT INTO CustomerCDP._TmpWAVerify_Verification_10102026
        (VerificationID, CustomerProfileID, WhatsAppNo, OldStatus, NewStatus, OldLastModifiedBy, OldLastModifiedTime, NeedUpdate)
    SELECT v.ID, v.CustomerProfileID, v.WhatsAppNo,
           ISNULL(CAST(v.WAVerifiedStatus AS VARCHAR(20)), ''),
           @FalseStatus,
           v.LastModifiedBy, v.LastModifiedTime,
           CASE WHEN ISNULL(CAST(v.WAVerifiedStatus AS VARCHAR(20)), '') <> @FalseStatus THEN 1 ELSE 0 END
    FROM CustomerCDP.Verification v
    JOIN #VerOwned vo ON v.WhatsAppNo = vo.WhatsAppNo
    WHERE v.WhatsAppNo IS NOT NULL AND LTRIM(RTRIM(v.WhatsAppNo)) <> ''
      AND NOT EXISTS (
          SELECT 1 FROM #Valid s WHERE s.CustomerProfileID = v.CustomerProfileID AND s.WhatsAppNo = v.WhatsAppNo
      );

    -- 4. Populate WhatsAppCDP log — OWNER rows (1)
    INSERT INTO CustomerCDP._TmpWAVerify_WhatsAppCDP_10102026
        (WhatsAppID, CustomerProfileID, WhatsAppNo, OldStatus, NewStatus, OldLastModifiedBy, OldLastModifiedTime, NeedUpdate)
    SELECT w.ID, w.CustomerProfileID, w.WhatsAppNo,
           ISNULL(CAST(w.WAVerifiedStaus AS INT), -1),
           1,
           w.LastModifiedBy, w.LastModifiedTime,
           CASE WHEN ISNULL(CAST(w.WAVerifiedStaus AS INT), -1) <> 1 THEN 1 ELSE 0 END
    FROM CustomerCDP.WhatsAppCDP w
    JOIN #Valid s ON w.CustomerProfileID = s.CustomerProfileID AND w.WhatsAppNo = s.WhatsAppNo;

    -- #WaOwned: nomor yang punya baris pemilik di WhatsAppCDP
    SELECT DISTINCT s.WhatsAppNo
    INTO #WaOwned
    FROM #Valid s
    WHERE EXISTS (
        SELECT 1 FROM CustomerCDP.WhatsAppCDP w
        WHERE w.CustomerProfileID = s.CustomerProfileID AND w.WhatsAppNo = s.WhatsAppNo
    );

    -- WhatsAppCDP log — OTHER rows (0), hanya kalau owner ada
    INSERT INTO CustomerCDP._TmpWAVerify_WhatsAppCDP_10102026
        (WhatsAppID, CustomerProfileID, WhatsAppNo, OldStatus, NewStatus, OldLastModifiedBy, OldLastModifiedTime, NeedUpdate)
    SELECT w.ID, w.CustomerProfileID, w.WhatsAppNo,
           ISNULL(CAST(w.WAVerifiedStaus AS INT), -1),
           0,
           w.LastModifiedBy, w.LastModifiedTime,
           CASE WHEN ISNULL(CAST(w.WAVerifiedStaus AS INT), -1) <> 0 THEN 1 ELSE 0 END
    FROM CustomerCDP.WhatsAppCDP w
    JOIN #WaOwned wo ON w.WhatsAppNo = wo.WhatsAppNo
    WHERE w.WhatsAppNo IS NOT NULL AND LTRIM(RTRIM(w.WhatsAppNo)) <> ''
      AND NOT EXISTS (
          SELECT 1 FROM #Valid s WHERE s.CustomerProfileID = w.CustomerProfileID AND s.WhatsAppNo = w.WhatsAppNo
      );

    -- 5. Summary
    DECLARE @VerTotal INT, @VerNeed INT, @VerOwner INT, @VerOther INT;
    DECLARE @WaTotal INT, @WaNeed INT, @WaOwner INT, @WaOther INT;

    SELECT @VerTotal = COUNT(*),
           @VerNeed  = SUM(CASE WHEN NeedUpdate = 1 THEN 1 ELSE 0 END),
           @VerOwner = SUM(CASE WHEN NewStatus = @TrueStatus THEN 1 ELSE 0 END),
           @VerOther = SUM(CASE WHEN NewStatus = @FalseStatus THEN 1 ELSE 0 END)
    FROM CustomerCDP._TmpWAVerify_Verification_10102026;

    SELECT @WaTotal = COUNT(*),
           @WaNeed  = SUM(CASE WHEN NeedUpdate = 1 THEN 1 ELSE 0 END),
           @WaOwner = SUM(CASE WHEN NewStatus = 1 THEN 1 ELSE 0 END),
           @WaOther = SUM(CASE WHEN NewStatus = 0 THEN 1 ELSE 0 END)
    FROM CustomerCDP._TmpWAVerify_WhatsAppCDP_10102026;

    RAISERROR('══════════════════════════════════════════════════', 0, 1) WITH NOWAIT;
    RAISERROR('  PREVIEW SUMMARY — WhatsApp Verification', 0, 1) WITH NOWAIT;
    RAISERROR('══════════════════════════════════════════════════', 0, 1) WITH NOWAIT;
    RAISERROR('  Source:    %d baris | Valid: %d | Konflik: %d', 0, 1, @TotalSrc, @ValidCnt, @ConflictCnt) WITH NOWAIT;
    RAISERROR('  Verification: %d di-log (%d owner, %d other) | %d perlu update', 0, 1, @VerTotal, @VerOwner, @VerOther, @VerNeed) WITH NOWAIT;
    RAISERROR('  WhatsAppCDP: %d di-log (%d owner, %d other) | %d perlu update', 0, 1, @WaTotal, @WaOwner, @WaOther, @WaNeed) WITH NOWAIT;
    RAISERROR('══════════════════════════════════════════════════', 0, 1) WITH NOWAIT;
    RAISERROR('  Cek tabel log. Kalau OK, ganti @Mode = EXECUTE dan jalankan lagi.', 0, 1) WITH NOWAIT;
    RAISERROR('══════════════════════════════════════════════════', 0, 1) WITH NOWAIT;
END

------------------------------------------------------------
-- EXECUTE: update data asli per batch
------------------------------------------------------------
IF @Mode = 'EXECUTE'
BEGIN
    IF OBJECT_ID('CustomerCDP._TmpWAVerify_Verification_10102026','U') IS NULL
       OR OBJECT_ID('CustomerCDP._TmpWAVerify_WhatsAppCDP_10102026','U') IS NULL
        THROW 50003, 'Tabel log tidak ada. Jalankan PREVIEW dulu.', 1;

    -- 1. EXECUTE: Verification
    SELECT @Max = ISNULL(MAX(RowNum), 0) FROM CustomerCDP._TmpWAVerify_Verification_10102026;
    SET @From = 1; SET @Total = 0;
    WHILE @From <= @Max
    BEGIN
        SET @To = @From + @BatchSize - 1;
        BEGIN TRY
            BEGIN TRAN;

            UPDATE v
            SET v.WAVerifiedStatus = l.NewStatus,
                v.LastModifiedBy   = @By,
                v.LastModifiedTime = @Today
            FROM CustomerCDP.Verification v
            JOIN CustomerCDP._TmpWAVerify_Verification_10102026 l ON l.VerificationID = v.ID
            WHERE l.RowNum BETWEEN @From AND @To
              AND l.NeedUpdate = 1
              AND l.IsUpdated = 0;

            SET @n = @@ROWCOUNT;

            UPDATE CustomerCDP._TmpWAVerify_Verification_10102026
            SET IsUpdated = 1, UpdatedTime = @Today
            WHERE RowNum BETWEEN @From AND @To AND NeedUpdate = 1 AND IsUpdated = 0;

            COMMIT TRAN;
        END TRY
        BEGIN CATCH
            IF XACT_STATE() <> 0 ROLLBACK TRAN;
            SET @Msg = CONVERT(VARCHAR(30), GETDATE(), 120) + ' - ERROR Verification RowNum ' + CONVERT(VARCHAR(20), @From) + ': ' + ERROR_MESSAGE();
            RAISERROR(@Msg, 0, 1) WITH NOWAIT;
            THROW;
        END CATCH

        SET @Total += @n;
        SET @Msg = CONVERT(VARCHAR(30), GETDATE(), 120) + ' - Verification RowNum ' + CONVERT(VARCHAR(20), @From)
                 + ' s/d ' + CONVERT(VARCHAR(20), CASE WHEN @To > @Max THEN @Max ELSE @To END) + ' dari ' + CONVERT(VARCHAR(20), @Max)
                 + ' | ter-update batch ini: ' + CONVERT(VARCHAR(20), @n) + ' | total: ' + CONVERT(VARCHAR(20), @Total);
        RAISERROR(@Msg, 0, 1) WITH NOWAIT;
        SET @From = @To + 1;
    END

    -- 2. EXECUTE: WhatsAppCDP
    SELECT @Max = ISNULL(MAX(RowNum), 0) FROM CustomerCDP._TmpWAVerify_WhatsAppCDP_10102026;
    SET @From = 1; SET @Total = 0;
    WHILE @From <= @Max
    BEGIN
        SET @To = @From + @BatchSize - 1;
        BEGIN TRY
            BEGIN TRAN;

            UPDATE w
            SET w.WAVerifiedStaus = l.NewStatus,
                w.LastModifiedBy   = @By,
                w.LastModifiedTime = @Today
            FROM CustomerCDP.WhatsAppCDP w
            JOIN CustomerCDP._TmpWAVerify_WhatsAppCDP_10102026 l ON l.WhatsAppID = w.ID
            WHERE l.RowNum BETWEEN @From AND @To
              AND l.NeedUpdate = 1
              AND l.IsUpdated = 0;

            SET @n = @@ROWCOUNT;

            UPDATE CustomerCDP._TmpWAVerify_WhatsAppCDP_10102026
            SET IsUpdated = 1, UpdatedTime = @Today
            WHERE RowNum BETWEEN @From AND @To AND NeedUpdate = 1 AND IsUpdated = 0;

            COMMIT TRAN;
        END TRY
        BEGIN CATCH
            IF XACT_STATE() <> 0 ROLLBACK TRAN;
            SET @Msg = CONVERT(VARCHAR(30), GETDATE(), 120) + ' - ERROR WhatsAppCDP RowNum ' + CONVERT(VARCHAR(20), @From) + ': ' + ERROR_MESSAGE();
            RAISERROR(@Msg, 0, 1) WITH NOWAIT;
            THROW;
        END CATCH

        SET @Total += @n;
        SET @Msg = CONVERT(VARCHAR(30), GETDATE(), 120) + ' - WhatsAppCDP RowNum ' + CONVERT(VARCHAR(20), @From)
                 + ' s/d ' + CONVERT(VARCHAR(20), CASE WHEN @To > @Max THEN @Max ELSE @To END) + ' dari ' + CONVERT(VARCHAR(20), @Max)
                 + ' | ter-update batch ini: ' + CONVERT(VARCHAR(20), @n) + ' | total: ' + CONVERT(VARCHAR(20), @Total);
        RAISERROR(@Msg, 0, 1) WITH NOWAIT;
        SET @From = @To + 1;
    END

    RAISERROR('EXECUTE SELESAI.', 0, 1) WITH NOWAIT;
END

-- Cleanup temp tables
IF OBJECT_ID('tempdb..#Src')      IS NOT NULL DROP TABLE #Src;
IF OBJECT_ID('tempdb..#Valid')    IS NOT NULL DROP TABLE #Valid;
IF OBJECT_ID('tempdb..#VerOwned') IS NOT NULL DROP TABLE #VerOwned;
IF OBJECT_ID('tempdb..#WaOwned')  IS NOT NULL DROP TABLE #WaOwned;

def reserve(db, user, seat, day, start, end):
    db.execute("BEGIN IMMEDIATE")
    try:
        if not db.execute("SELECT 1 FROM users WHERE id=?", (user,)).fetchone():
            raise ValueError("unknown-user")
        if not db.execute("SELECT 1 FROM seats WHERE id=?", (seat,)).fetchone():
            raise ValueError("unknown-seat")
        overlap = "day=? AND status='active' AND start<? AND end>?"
        query = "SELECT 1 FROM bookings WHERE "
        if db.execute(query + "seat_id=? AND " + overlap,
                      (seat, day, end, start)).fetchone():
            raise ValueError("seat-conflict")
        if db.execute(query + "user_id=? AND " + overlap,
                      (user, day, end, start)).fetchone():
            raise ValueError("user-conflict")
        cursor = db.execute(
            "INSERT INTO bookings(user_id,seat_id,day,start,end) "
            "VALUES (?,?,?,?,?)", (user, seat, day, start, end))
        db.execute("COMMIT")
        return "created", cursor.lastrowid
    except ValueError as error:
        db.execute("ROLLBACK")
        return str(error), None
    except BaseException:
        db.execute("ROLLBACK")
        raise

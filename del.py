

import sqlite3

# Function to delete booking by ID
def delete_booking_by_id(booking_id):
    conn = sqlite3.connect('slot_booking_new.db')
    c = conn.cursor()

    # Execute delete operation
    c.execute("DELETE FROM appointment_bookings WHERE id = ?", (booking_id,))

    # Commit changes and close connection
    conn.commit()
    conn.close()

    print(f"Booking with ID {booking_id} deleted successfully.")

# Example usage
if __name__ == '__main__':
    booking_id = input("Enter booking ID to delete: ")
    delete_booking_by_id(booking_id)

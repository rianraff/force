import os
import psycopg2
from dotenv import load_dotenv
from flask import Flask, request, jsonify

#----------------------
# Force Database
#----------------------
    
# Load environment variables from .env file
load_dotenv()

# Create a PostgreSQL connection
conn = psycopg2.connect(
    dbname=os.getenv("POSTGRES_DB"),
    user=os.getenv("POSTGRES_USER"),
    password=os.getenv("POSTGRES_PASSWORD"),
    host=os.getenv("POSTGRES_HOST"),
    port=os.getenv("POSTGRES_PORT")
)

# Initialize Flask app
app = Flask(__name__)

#-----------------------------------
# Route to insert data into the database
#-----------------------------------
# Route to insert data into the database
@app.route('/insert_data', methods=['POST'])
def insert_data():
    data = request.get_json()
    cluster_id = data.get('cluster_id')
    processed = data.get('processed')
    file_path = data.get('file_path')
    revise = data.get('revise')

    if not cluster_id or not processed or not revise:
        return jsonify({'message': 'Missing required parameters'}), 400

    insert_data_to_db(conn, cluster_id, processed, file_path, revise)
    return jsonify({'message': 'Data inserted successfully'}), 200

def insert_data_to_db(conn, cluster_id, processed, file_path, revise):
    # Create a new cursor
    cur = conn.cursor()

    # Execute the INSERT statement
    cur.execute(
        "INSERT INTO logging (cluster_id, processed, file_path, revise) VALUES (%s, %s, %s, %s)",
        (cluster_id, processed, file_path, revise)
    )

    # Commit the changes to the database
    conn.commit()

    # Close the cursor
    cur.close()

#-----------------------------------
# Route to get list of clusters
#-----------------------------------
# Route to get list of clusters
@app.route('/clusters', methods=['GET'])
def clusters():
    return getClusterList()

def getClusterList():
    try:
        # Create a new cursor
        cur = conn.cursor()

        # Execute a query to select all clusters from the logging table
        cur.execute("SELECT DISTINCT cluster_id FROM logging")

        # Fetch all rows from the result set
        rows = cur.fetchall()

        # Close the cursor
        cur.close()

        # Return the list of clusters as a JSON response
        return jsonify({"clusters": [row[0] for row in rows]})

    except psycopg2.Error as e:
        # Handle any errors that occur during database operation
        print(f"Error retrieving cluster list: {e}")
        return jsonify({"error": "Failed to retrieve cluster list"}), 500

#-----------------------------------
# Route to get data by cluster_id
#-----------------------------------
@app.route('/update_path', methods=['POST'])
def update_path():
    data = request.json
    cluster_id = data.get('cluster_id')
    file_path = data.get('file_path')

    if not cluster_id or not file_path:
        return jsonify({'error': 'Missing cluster_id or file_path parameter'}), 400

    update_path_in_db(conn, cluster_id, file_path)
    return jsonify({'message': 'Path updated successfully'}), 200

def update_path_in_db(conn, cluster_id, file_path):
    # Create a new cursor
    cur = conn.cursor()

    # Execute the UPDATE statement
    cur.execute(
        "UPDATE logging SET file_path = %s WHERE cluster_id = %s",
        (file_path, cluster_id)
    )

    # Commit the changes to the database
    conn.commit()

    # Close the cursor
    cur.close()

#-----------------------------------
# Route to get data by cluster_id
#-----------------------------------
@app.route('/get_path_cluster_id', methods=['GET'])
def get_path_cluster_id():
    data = get_data_from_db()
    if data:
        file_path, cluster_id = data
        result = {'file_path': file_path, 'cluster_id': cluster_id}
    else:
        result = {'file_path': None, 'cluster_id': None}
    return jsonify(result)

def get_data_from_db():
    cur = conn.cursor()
    cur.execute(
        "SELECT file_path, cluster_id FROM logging WHERE processed = 'FALSE' OR (processed = 'TRUE' AND revise = 'TRUE') ORDER BY id DESC LIMIT 1"
    )
    row = cur.fetchone()
    cur.close()
    return row

@app.route('/update_processed', methods=['POST'])
def update_processed():
    data = request.json
    cluster_id = data.get('cluster_id')

    if not cluster_id:
        return jsonify({'error': 'Missing cluster_id parameter'}), 400

    update_processed_in_db(conn, cluster_id)
    return jsonify({'message': 'Processed updated successfully'}), 200

def update_processed_in_db(conn, cluster_id):
    # Create a new cursor
    cur = conn.cursor()

    # Execute the UPDATE statement
    cur.execute(
        "UPDATE logging SET processed = 'TRUE' WHERE cluster_id = %s", (cluster_id,)
    )

    # Commit the changes to the database
    conn.commit()

    # Close the cursor
    cur.close()

@app.route('/update_revise', methods=['PUT'])
def update_revise():
    data = request.json
    cluster_id = data.get('cluster_id')
    revise = data.get('revise')

    if not cluster_id:
        return jsonify({'error': 'Missing cluster_id parameter'}), 400

    update_revise_in_db(conn, cluster_id, revise)
    return jsonify({'message': 'Processed updated successfully'}), 200

def update_revise_in_db(conn, cluster_id, revise):
    # Create a new cursor
    cur = conn.cursor()

    # Execute the UPDATE stateme\nt
    cur.execute(
        "UPDATE logging SET revise = %s WHERE cluster_id = %s", (revise, cluster_id,)
    )

    # Commit the changes to the database
    conn.commit()

    # Close the cursor
    cur.close()

#-----------------------------------
# Route to get all cluster ids
#-----------------------------------
@app.route('/get_all_cluster_ids', methods=['GET'])
def get_all_cluster_ids():
    # Membuat koneksi ke database
    conn = psycopg2.connect(
        dbname=os.getenv("POSTGRES_DB"),
        user=os.getenv("POSTGRES_USER"),
        password=os.getenv("POSTGRES_PASSWORD"),
        host=os.getenv("POSTGRES_HOST"),
        port=os.getenv("POSTGRES_PORT")
    )

    cur = conn.cursor()
    cur.execute("SELECT DISTINCT cluster_id FROM logging WHERE revise != 'TRUE'")
    rows = cur.fetchall()
    cur.close()
    conn.close()

    # Mengembalikan cluster_id sebagai list
    cluster_ids = [row[0] for row in rows]
    return jsonify(cluster_ids)

#-----------------------------------
# Main function to run the Flask app
#-----------------------------------
def main():
    app.run(debug=True)

if __name__ == '__main__':
    main()
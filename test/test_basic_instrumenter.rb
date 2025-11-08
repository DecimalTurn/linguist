require_relative "./helper"

class TestBasicInstrumenter < Minitest::Test
  include Linguist

  def setup
    @instrumenter = Linguist::BasicInstrumenter.new
    Linguist.instrumenter = @instrumenter
  end

  def teardown
    Linguist.instrumenter = nil
  end

  def test_tracks_extension_strategy
    # Ruby file detected by extension
    blob = fixture_blob("Ruby/foo.rb")
    Linguist.detect(blob)

    assert @instrumenter.detected_info.key?(blob.name)
    assert_equal "Extension", @instrumenter.detected_info[blob.name][:strategy]
    assert_equal "Ruby", @instrumenter.detected_info[blob.name][:language]
  end

  def test_tracks_modeline_strategy
    # File with vim modeline
    blob = fixture_blob("Data/Modelines/ruby")
    Linguist.detect(blob)

    assert @instrumenter.detected_info.key?(blob.name)
    assert_equal "Modeline", @instrumenter.detected_info[blob.name][:strategy]
    assert_equal "Ruby", @instrumenter.detected_info[blob.name][:language]
  end

  def test_tracks_shebang_strategy
    # File with shebang
    blob = fixture_blob("Shell/sh")
    Linguist.detect(blob)

    assert @instrumenter.detected_info.key?(blob.name)
    assert_equal "Shebang", @instrumenter.detected_info[blob.name][:strategy]
    assert_equal "Shell", @instrumenter.detected_info[blob.name][:language]
  end

  def test_tracks_multiple_files
    # Track multiple files in sequence
    ruby_blob = fixture_blob("Ruby/foo.rb")
    shell_blob = fixture_blob("Shell/sh")

    Linguist.detect(ruby_blob)
    Linguist.detect(shell_blob)

    assert_equal 2, @instrumenter.detected_info.size
    assert @instrumenter.detected_info.key?(ruby_blob.name)
    assert @instrumenter.detected_info.key?(shell_blob.name)
  end

  def test_no_tracking_for_binary_files
    binary_blob = fixture_blob("Binary/octocat.ai")
    Linguist.detect(binary_blob)

    # Should not record info for binary files
    assert_equal 0, @instrumenter.detected_info.size
  end

  def test_records_correct_strategy_for_heuristics
    # .bas file that should be detected via heuristics
    blob = fixture_blob("VBA/sample.bas")
    Linguist.detect(blob)

    assert @instrumenter.detected_info.key?(blob.name)
    # Multiple strategies may contribute, but heuristics should be involved
    strategy_chain = @instrumenter.detected_info[blob.name][:strategy]
    assert strategy_chain.include?("Heuristics"), "Expected Heuristics to be part of strategy chain: #{strategy_chain}"
  end

  def test_tracks_filename_strategy
    # Dockerfile detected by filename
    blob = fixture_blob("Dockerfile/Dockerfile")
    Linguist.detect(blob)

    assert @instrumenter.detected_info.key?(blob.name)
    assert_equal "Filename", @instrumenter.detected_info[blob.name][:strategy]
    assert_equal "Dockerfile", @instrumenter.detected_info[blob.name][:language]
  end

  def test_tracks_override_strategy
    # Simulate a blob with a gitattributes override
    blob = Linguist::FileBlob.new("Gemfile", "")
    # Simulate detection with gitattributes strategy showing the override
    strategy = Struct.new(:name).new("Filename (overridden by .gitattributes)")
    language = Struct.new(:name).new("Java")
    @instrumenter.instrument("linguist.detected", blob: blob, strategy: strategy, language: language) {}
    assert @instrumenter.detected_info.key?(blob.name)
    assert_match(/overridden by \.gitattributes/, @instrumenter.detected_info[blob.name][:strategy])
    assert_equal "Java", @instrumenter.detected_info[blob.name][:language]
  end

  def test_tracks_combined_strategies
    # .pl file that requires both Extension strategy (to identify .pl candidates)
    # and Heuristics strategy (to resolve between Perl and Raku)
    blob = fixture_blob("Raku/chromosome.pl")
    Linguist.detect(blob)

    assert @instrumenter.detected_info.key?(blob.name)
    strategy_chain = @instrumenter.detected_info[blob.name][:strategy]

    # Should contain both Extension (provides .pl candidates) and Heuristics (resolves to Raku)
    assert strategy_chain.include?("Extension"), "Expected Extension to be part of strategy chain: #{strategy_chain}"
    assert strategy_chain.include?("Heuristics"), "Expected Heuristics to be part of strategy chain: #{strategy_chain}"
    assert_equal "Raku", @instrumenter.detected_info[blob.name][:language]

    # The combined strategy should show both contributing strategies
    assert_match(/Extension.*Heuristics/, strategy_chain)
  end
end

def test_override_strategy_is_recorded
  # This file is overridden by .gitattributes to be detectable and language Markdown
  blob = sample_blob("Markdown/tender.md")
  Linguist.detect(blob)
  assert @instrumenter.detected_info.key?(blob.name)
  assert_includes ["GitAttributes"], @instrumenter.detected_info[blob.name][:strategy]
  assert_equal "Markdown", @instrumenter.detected_info[blob.name][:language]
end

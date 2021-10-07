{-# LANGUAGE CPP #-}
{-# LANGUAGE OverloadedStrings #-}
{-# LANGUAGE TemplateHaskell   #-}
{-# LANGUAGE RecordWildCards   #-}
{-# LANGUAGE DeriveGeneric #-}
module Codec.Xlsx.Types.HeaderFooter where

#ifdef USE_MICROLENS
import Lens.Micro.TH (makeLenses)
#else
import Control.Lens (makeLenses)
#endif

import Control.DeepSeq (NFData)
import Data.Default
import qualified Data.Map as Map
import Data.Maybe (catMaybes, listToMaybe)
import Data.Text (Text)
import GHC.Generics (Generic)
import Text.XML
import Text.XML.Cursor

import Codec.Xlsx.Parser.Internal
import Codec.Xlsx.Writer.Internal

-- | See 18.3.1.46 "headerFooter (Header Footer Settings)"
data HeaderFooter = HeaderFooter {
    _headerFooterOddHeader        :: Maybe Text
  , _headerFooterOddFooter        :: Maybe Text
  , _headerFooterEvenHeader       :: Maybe Text
  , _headerFooterEvenFooter       :: Maybe Text
  , _headerFooterFirstHeader      :: Maybe Text
  , _headerFooterFirstFooter      :: Maybe Text
  , _headerFooterDifferentOddEven :: Maybe Bool
  , _headerFooterDifferentFirst   :: Maybe Bool
  , _headerFooterScaleWithDoc     :: Maybe Bool
  , _headerFooterAlignWithMargins :: Maybe Bool
  }
  deriving (Eq, Ord, Show, Generic)
instance NFData HeaderFooter

instance Default HeaderFooter where
  def = HeaderFooter {
      _headerFooterOddHeader        = Nothing
    , _headerFooterOddFooter        = Nothing
    , _headerFooterEvenHeader       = Nothing
    , _headerFooterEvenFooter       = Nothing
    , _headerFooterFirstHeader      = Nothing
    , _headerFooterFirstFooter      = Nothing
    , _headerFooterDifferentOddEven = Nothing
    , _headerFooterDifferentFirst   = Nothing
    , _headerFooterScaleWithDoc     = Nothing
    , _headerFooterAlignWithMargins = Nothing
    }

makeLenses ''HeaderFooter

instance ToElement HeaderFooter where
  toElement nm HeaderFooter{..} = Element {
      elementName       = nm
    , elementNodes      = map NodeElement . catMaybes $ [
          fmap (elementContent "oddHeader")   _headerFooterOddHeader
        , fmap (elementContent "oddFooter")   _headerFooterOddFooter
        , fmap (elementContent "evenHeader")  _headerFooterEvenHeader
        , fmap (elementContent "evenFooter")  _headerFooterEvenFooter
        , fmap (elementContent "firstHeader") _headerFooterFirstHeader
        , fmap (elementContent "firstFooter") _headerFooterFirstFooter
        ]
    , elementAttributes = Map.fromList . catMaybes $ [
          "differentOddEven" .=? _headerFooterDifferentOddEven
        , "differentFirst"   .=? _headerFooterDifferentFirst
        , "scaleWithDoc"     .=? _headerFooterScaleWithDoc
        , "alignWithMargins" .=? _headerFooterAlignWithMargins
        ]
    }

instance FromCursor HeaderFooter where
  fromCursor cur = do
    let
      _headerFooterOddHeader   = listToMaybe $ cur $/ element (n_ "oddHeader") &/ content
      _headerFooterOddFooter   = listToMaybe $ cur $/ element (n_ "oddFooter") &/ content
      _headerFooterEvenHeader  = listToMaybe $ cur $/ element (n_ "evenHeader") &/ content
      _headerFooterEvenFooter  = listToMaybe $ cur $/ element (n_ "evenFooter") &/ content
      _headerFooterFirstHeader = listToMaybe $ cur $/ element (n_ "firstHeader") &/ content
      _headerFooterFirstFooter = listToMaybe $ cur $/ element (n_ "firstFooter") &/ content
    _headerFooterDifferentOddEven <- maybeAttribute "differentOddEven" cur
    _headerFooterDifferentFirst   <- maybeAttribute "differentFirst" cur
    _headerFooterScaleWithDoc     <- maybeAttribute "scaleWithDoc" cur
    _headerFooterAlignWithMargins <- maybeAttribute "alignWithMargins" cur
    return HeaderFooter{..}

instance FromXenoNode HeaderFooter where
  fromXenoNode root = parseAttributes root $ do
    ( _headerFooterOddHeader
      , _headerFooterOddFooter
      , _headerFooterEvenHeader
      , _headerFooterEvenFooter
      , _headerFooterFirstHeader
      , _headerFooterFirstFooter
      ) <-
      toAttrParser . collectChildren root $
      (,,,,,) <$> maybeParse "oddHeader" contentX
              <*> maybeParse "oddFooter" contentX
              <*> maybeParse "evenHeader" contentX
              <*> maybeParse "evenFooter" contentX
              <*> maybeParse "firstHeader" contentX
              <*> maybeParse "firstFooter" contentX
    _headerFooterDifferentOddEven <- maybeAttr "differentOddEven"
    _headerFooterDifferentFirst   <- maybeAttr "differentFirst"
    _headerFooterScaleWithDoc     <- maybeAttr "scaleWithDoc"
    _headerFooterAlignWithMargins <- maybeAttr "alignWithMargins"
    return HeaderFooter{..}
